"""document_reader.py

Reads an existing .docx and exposes its paragraphs (body + table cells) in
document order. Deliberately does NOT recreate or modify the document; it only
inspects structure so that downstream modules can preserve it.
"""
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def is_heading(paragraph):
    """Heuristic: is this paragraph a section heading?

    Uses the paragraph's style first (Heading 1..9), then falls back to a
    bold + short heuristic for documents that only use direct formatting.
    """
    try:
        style = paragraph.style
        if style is not None and style.name and style.name.startswith("Heading"):
            return True
    except Exception:
        pass
    runs = [r for r in paragraph.runs if (r.text or "").strip()]
    if runs and all(r.bold for r in runs) and len(paragraph.text.strip()) < 90:
        return True
    return False


def iter_paragraphs(doc):
    """Yield python-docx Paragraph objects for body paragraphs and table-cell
    paragraphs, in document order. This is the single source of truth used by
    both the reader and the editor so indices always align."""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield from _iter_table(child, doc)


def _iter_table(tbl, doc):
    for tr in tbl.findall(qn("w:tr")):
        for tc in tr.findall(qn("w:tc")):
            for p in tc.findall(qn("w:p")):
                yield Paragraph(p, doc)


class DocumentReader:
    def __init__(self, path):
        self.path = str(path)
        self.document = Document(self.path)

    def all_paragraphs(self):
        return list(iter_paragraphs(self.document))

    def count_existing_footnotes(self):
        """Count existing Word footnote references anywhere in the document
        (body, tables, hyperlinks). Used for the 'existing preserved' check."""
        n = 0
        for p in self.all_paragraphs():
            n += len(p._p.findall(".//" + qn("w:footnoteReference")))
        return n
