"""agent_tools.py

Thin, narrow wrappers around the deterministic pipeline. An AI agent (future
stage) can call these tools instead of manipulating the document directly. Each
tool has a single, explicit responsibility and returns plain data - the agent
is never given unrestricted filesystem or shell control. It cannot create,
delete or rewrite arbitrary files: the only file operations available to it are
"read this .docx and report" (inspect_document) and "run the deterministic
pipeline on one .docx into the standard output path" (run_pipeline).

Future Mendeley stage can build on the same shape:
    citation_marker -> reference_url -> identify publication -> Mendeley ref
    -> Mendeley Cite citation -> updated bibliography
(currently stops at: citation_marker -> reference_url -> Word footnote)
"""
from config import Config
from document_reader import DocumentReader, is_heading
from reference_parser import ReferenceParser
from citation_detector import CitationDetector
from validator import file_hash, validate_output, ValidationContext
from report import Report


def inspect_document(path):
    """Return a high-level description of the document (counts, headings,
    existing footnotes). Read-only; no files are created or modified."""
    reader = DocumentReader(path)
    paras = reader.all_paragraphs()
    headings = [p.text for p in paras if is_heading(p)]
    snap = reader.existing_footnotes()
    return {
        "path": str(path),
        "sha256": file_hash(path),
        "paragraph_count": len(paras),
        "existing_footnotes": len(snap.definitions),
        "headings": headings[:50],
    }


def find_reference_section(paragraphs, config):
    """Locate the reference section. Returns (start_index, method, confidence,
    note) or None (-> caller must abort, no guessing)."""
    return ReferenceParser(config).find_reference_section(paragraphs, config)


def extract_references(paragraphs, start_index):
    """Extract the number -> {url, text} map from the reference list."""
    return ReferenceParser().parse_references(paragraphs, start_index)


def find_citations(paragraphs, ref_section_start):
    """Detect citation markers in the body (index < ref_section_start)."""
    return CitationDetector().detect(paragraphs, ref_section_start)


def validate_citation(citation, ref_map):
    """Return (ok, reason) for a citation against the reference map."""
    for num in citation.numbers:
        if num not in ref_map:
            return False, f"Reference {num} not found"
        if not ref_map[num]["url"]:
            return False, f"Reference {num} has no URL"
    return True, ""


def verify_output(ctx: ValidationContext):
    """Run the full post-save validation. Returns {"ok", "failures", "results"}."""
    return validate_output(ctx)


def generate_report(path, rows, summary):
    """Write the CSV review report at ``path`` (the only file an agent may
    write through these tools)."""
    rep = Report()
    for row in rows:
        rep.add(*row)
    rep.set_summary(**summary)
    rep.write(path)
    return path


def run_pipeline(input_path, output_path=None, backend="xml",
                 keep_marker=False, overwrite=False,
                 input_dir="input", output_dir="output"):
    """One-call deterministic pipeline (the only way an agent may produce a
    modified document). Returns (exit_code, output_path, report_path)."""
    from main import run  # local import: keep module import side-effect free

    config = Config()
    config.input_path = input_path
    config.output_path = output_path
    config.backend = backend
    config.replace_marker = not keep_marker
    config.overwrite_output = overwrite
    config.input_dir = input_dir
    config.output_dir = output_dir
    from pathlib import Path
    code = run(config)
    out = config.output_path or str(
        Path(output_dir) / (Path(input_path).stem + "_with_footnotes.docx"))
    return code, out, Path(out).with_name("citation_review_report.csv")
