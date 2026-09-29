"""agent_tools.py

Thin, narrow wrappers around the deterministic pipeline. An AI agent (future
stage) can call these tools instead of manipulating the document directly. Each
tool has a single, explicit responsibility and returns plain data - the agent is
never given unrestricted filesystem or shell control.

Future Mendeley stage can build on the same shape:
    citation_marker -> reference_url -> identify publication -> Mendeley ref
    -> Mendeley Cite citation -> updated bibliography
(currently stops at: citation_marker -> reference_url -> Word footnote)
"""
from document_reader import DocumentReader, is_heading, iter_paragraphs
from reference_parser import ReferenceParser
from citation_detector import CitationDetector, Citation
from word_footnotes import (
    XmlFootnoteBackend, process_paragraph, ensure_footnote_styles)
from validator import file_hash, original_unchanged, validate_output
from report import Report


def inspect_document(path):
    """Return a high-level description of the document (counts, headings)."""
    reader = DocumentReader(path)
    paras = reader.all_paragraphs()
    headings = [p.text for p in paras if is_heading(p)]
    return {
        "path": str(path),
        "paragraph_count": len(paras),
        "existing_footnotes": reader.count_existing_footnotes(),
        "headings": headings[:50],
    }


def find_reference_section(reader_paragraphs, config):
    parser = ReferenceParser()
    return parser.find_reference_section(reader_paragraphs, config)


def extract_references(parser, reader_paragraphs, start_index):
    return parser.parse_references(reader_paragraphs, start_index)


def find_citations(reader_paragraphs, ref_section_start):
    detector = CitationDetector()
    return detector.detect(reader_paragraphs, ref_section_start)


def validate_citation(citation, ref_map):
    """Return (ok, reason)."""
    for num in citation.numbers:
        if num not in ref_map:
            return False, f"Reference {num} not found"
        if not ref_map[num]["url"]:
            return False, f"Reference {num} has no URL"
    return True, ""


def insert_footnote():
    """Placeholder for the per-citation insertion primitive; the real work is done
    by word_footnotes.process_paragraph (called by the orchestrator). Kept as an
    explicit, named step for agent clarity."""
    raise NotImplementedError(
        "Per-citation insertion is performed by word_footnotes.process_paragraph "
        "during orchestration.")


def verify_output(out_path, ref_map, citations, existing_footnote_count):
    return validate_output(out_path, ref_map, citations, existing_footnote_count)


def generate_report(path, rows, summary):
    rep = Report()
    for r in rows:
        rep.add(*r)
    rep.set_summary(**summary)
    rep.write(path)
    return path
