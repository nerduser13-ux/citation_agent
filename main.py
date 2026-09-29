"""main.py

Orchestrates the workflow:

    input .docx  ->  hash  ->  read (read-only)  ->  locate reference section
    ->  extract number->URL map  ->  detect body citations  ->  insert REAL
    Word footnotes in a separate in-memory instance  ->  save to NEW output
    path  ->  validate  ->  output .docx + citation_review_report.csv

The input file is NEVER modified; its SHA-256 is verified before and after.

Run:
    python main.py                          # first .docx in ./input
    python main.py --input thesis.docx
    python main.py --input thesis.docx --output out.docx --keep-marker
    python main.py --backend xml            # default; --backend com is
                                            # experimental, Windows-only

Exit codes:
    0  success (output produced and validated)
    1  aborted (reference section not found) or validation failed
    2  usage error / unsupported backend
    3  input file changed during processing (fatal)
"""
import argparse
import sys
from collections import defaultdict
from pathlib import Path

from docx import Document

from config import Config
from document_reader import DocumentReader, is_heading, iter_paragraphs
from reference_parser import ReferenceParser
from citation_detector import CitationDetector
from word_footnotes import (XmlFootnoteBackend, ensure_footnote_styles,
                            process_document_com, process_paragraph)
from validator import (ValidationContext, file_hash, original_unchanged,
                       validate_output)
from report import Report


def _parse_args(argv):
    p = argparse.ArgumentParser(
        description="Insert real Word footnotes (source URLs) for [n] citations.")
    p.add_argument("--input", help="Input .docx (default: first file in ./input)")
    p.add_argument("--output", help="Output .docx (default: ./output/..._with_footnotes.docx)")
    p.add_argument("--backend", choices=["xml", "com"], default="xml",
                   help="Footnote backend (default: xml / OOXML)")
    p.add_argument("--keep-marker", action="store_true",
                   help="Keep the [n] marker text and insert the footnote after it")
    p.add_argument("--overwrite", action="store_true",
                   help="Overwrite output if it already exists")
    p.add_argument("--input-dir", default="input")
    p.add_argument("--output-dir", default="output")
    return p.parse_args(argv)


def _current_section(paragraphs, idx):
    for j in range(idx, -1, -1):
        if is_heading(paragraphs[j]):
            return (paragraphs[j].text or "").strip()[:60] or "(unnamed heading)"
    return "(document start)"


def _build_report_rows(citations, paragraphs, locations, ref_map):
    rows = []
    for c in citations:
        where = locations[c.para_index] if c.para_index < len(locations) else "?"
        loc = (f"para {c.para_index + 1} ({where}); "
               f"section: {_current_section(paragraphs, c.para_index)}")
        if c.status == "SUCCESS":
            base_note = "Range expanded" if c.is_range else "Inline citation"
            for num in c.numbers:
                entry = ref_map.get(num, {})
                rows.append((num, loc, c.marker_text,
                             entry.get("text", "")[:160], entry.get("url", ""),
                             "Footnote inserted", "SUCCESS", base_note))
        else:
            num = c.numbers[0] if c.numbers else ""
            entry = ref_map.get(num, {}) if num in ref_map else {}
            rows.append((num, loc, c.marker_text,
                         entry.get("text", "")[:160], entry.get("url", ""),
                         "Not processed", c.status, c.note))
    return rows


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
    try:
        out_path = config.resolve_output_path()
    except ValueError as e:
        print(f"ERROR: {e}")
        return 2

    in_hash = file_hash(inp)
    print("Word Citation Footnote Agent")
    print(f"Input : {inp}")
    print(f"Output: {out_path}")
    print("Analyzing document...")

    # ---- read-only snapshot of the input --------------------------------
    reader = DocumentReader(inp)
    paragraphs = reader.all_paragraphs()
    locations = reader.locations()
    raw_texts = reader.raw_texts()
    existing = reader.existing_footnotes()

    parser = ReferenceParser(config)
    ref_start = parser.find_reference_section(paragraphs, config)
    if ref_start is None:
        print("ERROR: Could not confidently locate the reference/bibliography "
              "section. Stopping safely - the input file is unchanged.")
        rep = Report()
        rep.set_summary(
            Input=str(inp),
            Status="ABORTED - reference section could not be confidently identified",
            Reference_section="NOT FOUND",
            Reference_section_method="none",
            Reference_section_confidence="none",
            References_detected=0, URLs_detected=0, Citation_markers_detected=0,
            Footnotes_inserted=0, Existing_footnotes_detected=len(existing.definitions),
            Existing_footnotes_preserved="n/a (not processed)",
            Unresolved_citations=0, Missing_URLs="n/a (not processed)",
            Ambiguous_citations=0, Uncited_references="n/a (not processed)",
            Previously_processed=0, Output_valid_docx="n/a (not processed)",
            Reopened_with_python_docx="n/a (not processed)",
            All_footnote_URLs_from_refs="n/a (not processed)",
            Input_unchanged=original_unchanged(inp, in_hash),
        )
        report_path = out_path.with_name("citation_review_report.csv")
        rep.write(str(report_path))
        print(f"Report : {report_path}")
        return 1

    start_idx, method, confidence, note = ref_start
    ref_map, warnings = parser.parse_references(paragraphs, start_idx)
    citations = CitationDetector().detect(paragraphs, start_idx)

    print(f"Reference section : FOUND ({method}, {confidence}) - {note}")
    for w in warnings:
        print(f"  warning: {w}")
    print(f"References detected: {len(ref_map)}")
    print(f"URLs detected     : {sum(1 for v in ref_map.values() if v['url'])}")
    print(f"Citation markers  : {len(citations)}")
    print("Processing...")

    # ---- editing: a SEPARATE Document instance, saved to the new path ----
    new_footnotes = {}
    if config.backend == "com":
        # Windows Word COM (experimental). Operates on the input via Word and
        # SaveAs-es to the output path; the input file is never written.
        try:
            process_document_com(str(inp), citations, ref_map,
                                 config.replace_marker, str(out_path))
        except (RuntimeError, ImportError) as e:
            print(f"ERROR: COM backend unavailable: {e}")
            return 2
        # The fids Word allocated are unknown to us; diff output vs input to
        # discover the new footnotes (their content is the URL itself).
        from document_reader import snapshot_existing_footnotes
        out_snap = snapshot_existing_footnotes(Document(out_path))
        new_footnotes = {int(fid): text for fid, text in out_snap.definitions.items()
                         if fid not in existing.definitions}
    else:
        edit_doc = Document(inp)
        style_ids = ensure_footnote_styles(edit_doc)
        backend = XmlFootnoteBackend(edit_doc, style_ids)
        edit_paragraphs = list(iter_paragraphs(edit_doc))
        by_para = defaultdict(list)
        for c in citations:
            by_para[c.para_index].append(c)
        for idx, cites in by_para.items():
            process_paragraph(edit_paragraphs[idx]._p, cites, backend,
                              config.replace_marker, ref_map)
        new_footnotes = dict(backend.new_footnotes)
        backend.save()
        edit_doc.save(str(out_path))

    # ---- the input must be byte-identical at this point ------------------
    if not original_unchanged(inp, in_hash):
        print("FATAL: the input file was modified during processing. "
              "This must not happen; aborting with a failure.")
        return 3

    # ---- automated validation --------------------------------------------
    val = validate_output(ValidationContext(
        in_path=str(inp), out_path=str(out_path), in_hash=in_hash,
        ref_map=ref_map, citations=citations, replace_marker=config.replace_marker,
        existing=existing, input_raw_texts=raw_texts,
        new_footnotes=new_footnotes, parser_warnings=warnings))
    for w in val["failures"]:
        print(f"VALIDATION FAILURE: {w}")

    # ---- CSV review report -------------------------------------------------
    footnotes_inserted = len(new_footnotes)
    unresolved = [c for c in citations if c.status == "ERROR"]
    ambiguous = [c for c in citations if c.status == "AMBIGUOUS"]
    previously = [c for c in citations if c.status == "SKIPPED"]
    missing_urls = sorted({n for n, v in ref_map.items() if not v["url"]})
    cited = {n for c in citations if c.status in ("SUCCESS", "SKIPPED") for n in c.numbers}
    uncited = sorted(set(ref_map) - cited)

    rep = Report()
    for row in _build_report_rows(citations, paragraphs, locations, ref_map):
        rep.add(*row)
    r = val["results"]
    rep.set_summary(
        Input=str(inp),
        Output=str(out_path),
        Reference_section=f"FOUND ({method}, {confidence})",
        Reference_section_method=method,
        Reference_section_confidence=confidence,
        References_detected=len(ref_map),
        URLs_detected=sum(1 for v in ref_map.values() if v["url"]),
        Citation_markers_detected=len(citations),
        Footnotes_inserted=footnotes_inserted,
        Existing_footnotes_detected=len(existing.definitions),
        Existing_footnotes_preserved=r.get("existing_preserved", False),
        Unresolved_citations=len(unresolved),
        Missing_URLs=", ".join(str(x) for x in missing_urls) or "0",
        Ambiguous_citations=len(ambiguous),
        Uncited_references=", ".join(str(x) for x in uncited) or "0",
        Previously_processed=len(previously),
        Output_valid_docx=r.get("valid_zip", False),
        Reopened_with_python_docx=r.get("reopen_ok", False),
        All_footnote_URLs_from_refs=r.get("all_new_urls_from_refs", False),
        Input_unchanged=r.get("input_unchanged", False),
        Status="OK" if val["ok"] else "VALIDATION FAILED",
    )
    report_path = out_path.with_name("citation_review_report.csv")
    rep.write(str(report_path))

    print()
    print(f"Footnotes inserted : {footnotes_inserted}")
    print(f"Unresolved citations: {len(unresolved)}")
    print(f"Missing URLs       : {len(missing_urls)}")
    print(f"Ambiguous citations: {len(ambiguous)}")
    print(f"Validation         : {'PASS' if val['ok'] else 'FAIL'}")
    print()
    print(f"Output : {out_path}")
    print(f"Report : {report_path}")
    return 0 if val["ok"] else 1


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
