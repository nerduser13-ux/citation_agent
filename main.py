"""main.py

Orchestrates the workflow:

    input .docx  ->  copy/original untouched  ->  analyse  ->  insert footnotes
    ->  validate  ->  output .docx  +  citation_review_report.csv

Run:

    python main.py                 # uses the first .docx in ./input
    python main.py --input x.docx  # explicit file
    python main.py --input x.docx --output y.docx --keep-marker --backend xml

The input file is NEVER modified.
"""
import argparse
import sys
from collections import defaultdict
from pathlib import Path

from docx.oxml.ns import qn

from config import Config
from document_reader import DocumentReader, is_heading, iter_paragraphs
from reference_parser import ReferenceParser
from citation_detector import CitationDetector
from word_footnotes import XmlFootnoteBackend, process_paragraph, ensure_footnote_styles
from validator import file_hash, original_unchanged, validate_output
from report import Report


def _parse_args(argv):
    p = argparse.ArgumentParser(
        description="Insert real Word footnotes (source URLs) for [n] citations.")
    p.add_argument("--input", help="Input .docx (default: first file in ./input)")
    p.add_argument("--output", help="Output .docx (default: ./output/..._with_footnotes.docx)")
    p.add_argument("--backend", choices=["xml", "com"], default="xml",
                   help="Footnote backend (default: xml / OOXML)")
    p.add_argument("--keep-marker", action="store_true",
                   help="Keep the [n] marker text (default: replace it)")
    p.add_argument("--overwrite", action="store_true",
                   help="Overwrite output if it already exists")
    p.add_argument("--input-dir", default="input")
    p.add_argument("--output-dir", default="output")
    return p.parse_args(argv)


def _current_section(paragraphs, idx):
    for j in range(idx, -1, -1):
        if is_heading(paragraphs[j]):
            return paragraphs[j].text.strip()[:60]
    return "(document start)"


def run(config):
    if not config.input_path:
        docs = sorted(Path(config.input_dir).glob("*.docx"))
        if not docs:
            print(f"No .docx found in {config.input_dir}/. Provide --input.")
            return 2
        config.input_path = str(docs[0])

    inp = Path(config.input_path)
    if not inp.exists():
        print(f"Input not found: {inp}")
        return 2

    in_hash = file_hash(inp)
    out_path = config.resolve_output_path()

    print("Word Citation Footnote Agent")
    print(f"Input : {inp}")
    print(f"Output: {out_path}")
    print("Analyzing document...")

    reader = DocumentReader(inp)
    paragraphs = reader.all_paragraphs()

    parser = ReferenceParser()
    det = CitationDetector()

    ref_start = parser.find_reference_section(paragraphs, config)
    if ref_start is None:
        print("ERROR: Could not confidently locate the reference/bibliography "
              "section. Stopping safely - the input file is unchanged.")
        rep = Report()
        rep.set_summary(
            Input=str(inp), Reference_section="NOT FOUND",
            Status="ABORTED - reference section not identified")
        rep.write(out_path.with_name("citation_review_report.csv"))
        return 1

    start_idx, method, confidence, note = ref_start
    ref_map = parser.parse_references(paragraphs, start_idx)
    citations = det.detect(paragraphs, start_idx)

    # Validate citations against the reference map (pre-insertion).
    for c in citations:
        for num in c.numbers:
            if num not in ref_map:
                c.status = "ERROR"
                c.note = f"Reference {num} not found"
                break
            if not ref_map[num]["url"]:
                c.status = "ERROR"
                c.note = f"Reference {num} has no URL"
                break

    existing_footnotes = reader.count_existing_footnotes()

    print(f"Reference section : FOUND ({method}, {confidence})")
    print(f"References detected: {len(ref_map)}")
    print(f"URLs detected     : {sum(1 for v in ref_map.values() if v['url'])}")
    print(f"Citation markers  : {len(citations)}")
    print("Processing...")

    # Load the SAME input into a separate Document object for editing, then save
    # to out_path. The original file on disk is never written to.
    from docx import Document
    edit_doc = Document(inp)
    ensure_footnote_styles(edit_doc)

    if config.backend == "com":
        # Windows Word COM path (experimental). Operates on disk; we SaveAs to the
        # output path so the input stays untouched.
        from word_footnotes import process_document_com
        process_document_com(str(inp), citations, ref_map,
                             config.replace_marker, str(out_path))
        backend_inserted = sum(1 for c in citations if c.status == "SUCCESS")
    else:
        backend = XmlFootnoteBackend(edit_doc)
        edit_paragraphs = list(iter_paragraphs(edit_doc))

        # Idempotency guard: if keep-marker mode and the paragraph already holds a
        # footnote reference (i.e. this output was processed before), skip it.
        by_para = defaultdict(list)
        for c in citations:
            by_para[c.para_index].append(c)

        for idx, cites in by_para.items():
            p_elem = edit_paragraphs[idx]._p
            process_paragraph(p_elem, cites, backend, config.replace_marker, ref_map)

        backend.save()
        edit_doc.save(str(out_path))
        backend_inserted = backend.inserted

    # ---- Validation -------------------------------------------------------
    assert original_unchanged(inp, in_hash), "INPUT FILE WAS MODIFIED!"
    val = validate_output(str(out_path), ref_map, citations, existing_footnotes)

    # ---- Report -----------------------------------------------------------
    rep = Report()
    for c in citations:
        loc = f"Paragraph {c.para_index + 1} ({_current_section(paragraphs, c.para_index)})"
        if c.status == "SUCCESS":
            for num in c.numbers:
                rep.add(num, loc, c.marker_text,
                        ref_map.get(num, {}).get("text", "")[:160],
                        ref_map.get(num, {}).get("url", ""), "Footnote inserted",
                        "SUCCESS", "Range expanded" if c.is_range else "Inline citation")
        else:
            num = c.numbers[0] if c.numbers else ""
            rep.add(num, loc, c.marker_text,
                    ref_map.get(num, {}).get("text", "")[:160] if num in ref_map else "",
                    ref_map.get(num, {}).get("url", "") if num in ref_map else "",
                    "Not processed", c.status, c.note)

    unresolved = [c for c in citations if c.status == "ERROR"]
    ambiguous = [c for c in citations if "spans runs" in c.note or "drawing" in c.note]
    missing_urls = sorted({n for n, v in ref_map.items() if not v["url"]})
    cited = {n for c in citations if c.status == "SUCCESS" for n in c.numbers}
    uncited = sorted(set(ref_map) - cited)
    previously = [c for c in citations if c.status == "SKIPPED"]

    rep.set_summary(
        References_detected=len(ref_map),
        URLs_detected=sum(1 for v in ref_map.values() if v["url"]),
        Citation_markers_detected=len(citations),
        Footnotes_inserted=backend_inserted,
        Existing_footnotes_detected=existing_footnotes,
        Existing_footnotes_preserved=val.get("existing_preserved", False),
        Unresolved_citations=len(unresolved),
        Missing_URLs=", ".join(str(x) for x in missing_urls) or "0",
        Ambiguous_citations=len(ambiguous),
        Uncited_references=", ".join(str(x) for x in uncited) or "0",
        Previously_processed=len(previously),
        Output_valid_docx=val.get("valid_docx", False),
        Reopened_with_python_docx=val.get("reopen_ok", False),
        All_footnote_URLs_from_refs=val.get("all_footnote_urls_in_refs", False),
        Input_unchanged=original_unchanged(inp, in_hash),
    )
    report_path = out_path.with_name("citation_review_report.csv")
    rep.write(str(report_path))

    print()
    print(f"Footnotes inserted : {backend_inserted}")
    print(f"Unresolved citations: {len(unresolved)}")
    print(f"Missing URLs       : {len(missing_urls)}")
    print(f"Ambiguous citations: {len(ambiguous)}")
    print()
    print(f"Output : {out_path}")
    print(f"Report : {report_path}")
    return 0


def main(argv=None):
    args = _parse_args(sys.argv[1:] if argv is None else argv)
    config = Config()
    config.input_path = args.input
    config.output_path = args.output
    config.backend = args.backend
    config.replace_marker = not args.keep_marker
    config.overwrite_output = args.overwrite
    config.input_dir = Path(args.input_dir)
    config.output_dir = Path(args.output_dir)
    return run(config)


if __name__ == "__main__":
    raise SystemExit(main())
