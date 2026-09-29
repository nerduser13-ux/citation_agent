"""document_reader.py

Reads an existing .docx and exposes its paragraphs (body, table cells, and
content-control blocks) in document order. Deliberately does NOT recreate or
modify the document; it only inspects structure so that downstream modules can
preserve it.

The paragraph order produced by :func:`iter_paragraphs` is the SINGLE source of
truth shared by detection, processing, reporting and validation - a citation
found in table cell (row 2, col 1) maps to exactly the same ``<w:p>`` element
when editing.
"""
from dataclasses import dataclass
from typing import Iterator, List, Optional, Tuple

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# Element tags that may contain block-level content we must walk.
_P = qn("w:p")
_TBL = qn("w:tbl")
_TR = qn("w:tr")
_TC = qn("w:tc")
_SDT = qn("w:sdt")
_SDT_CONTENT = qn("w:sdtContent")


# --------------------------------------------------------------------------
# Paragraph traversal (single source of truth for ordering/indexing)
# --------------------------------------------------------------------------
def _iter_block_content(container, doc, where: str) -> Iterator[Tuple[Paragraph, str]]:
    """Yield (Paragraph, location) for block content inside ``container``.

    ``container`` is a ``<w:body>`` or a ``<w:sdt>`` content region. Recurses
    into nested tables and content controls so that detection/processing/
    reporting/validation all agree on indices.
    """
    for child in container:
        if child.tag == _P:
            yield Paragraph(child, doc), where
        elif child.tag == _TBL:
            yield from _iter_table(child, doc, where)
        elif child.tag == _SDT:
            yield from _iter_sdt(child, doc, where)
        # Everything else (w:sectPr, w:bookmarkStart, ...) is structural and
        # contains no paragraphs; leave it alone.


def _iter_table(tbl, doc, where: str) -> Iterator[Tuple[Paragraph, str]]:
    row = 0
    for tr in tbl.iterchildren(_TR):
        row += 1
        col = 0
        for tc in tr.iterchildren(_TC):
            col += 1
            cell_where = f"{where} > table r{row}c{col}" if where else f"table r{row}c{col}"
            yield from _iter_block_content(tc, doc, cell_where)


def _iter_sdt(sdt, doc, where: str) -> Iterator[Tuple[Paragraph, str]]:
    """A block-level content control is ``<w:sdt><w:sdtPr/><w:sdtEndPr/>
    <w:sdtContent>...block content...</w:sdtContent></w:sdt>``: its paragraphs
    live inside ``w:sdtContent``, never directly under ``w:sdt``."""
    sdt_where = f"{where} > sdt" if where else "sdt"
    content = sdt.find(_SDT_CONTENT)
    if content is None:
        return
    yield from _iter_block_content(content, doc, sdt_where)


def iter_paragraph_info(doc) -> Iterator[Tuple[Paragraph, str]]:
    """Yield (Paragraph, location) for every block paragraph in document order.

    ``location`` is a human-readable descriptor ("" for direct body paragraphs,
    "table r1c2" for a cell, "sdt" for a content-control block, combinable).
    """
    yield from _iter_block_content(doc.element.body, doc, "")


def iter_paragraphs(doc) -> Iterator[Paragraph]:
    """Yield python-docx Paragraph objects in document order (body + tables +
    content controls). Used as the single coordinate system everywhere."""
    for p, _where in iter_paragraph_info(doc):
        yield p


def paragraph_locations(doc) -> List[str]:
    """Aligned location descriptors for :func:`iter_paragraphs` output."""
    return [where or "body" for _p, where in iter_paragraph_info(doc)]


def in_table(paragraph) -> bool:
    """True when the paragraph sits inside a table cell (at any depth)."""
    return any(a.tag == _TC for a in paragraph._p.iterancestors())


# --------------------------------------------------------------------------
# Raw text helpers
# --------------------------------------------------------------------------
def raw_text_of(p_elem) -> str:
    """Concatenated text of every ``<w:t>`` under ``p_elem`` in document order
    (includes runs inside hyperlinks and inline content controls; excludes
    footnote-reference runs, which carry no text).

    This is the ground-truth text used for fidelity checking; it never relies
    on ``Paragraph.text`` (which skips some structural positions)."""
    return "".join(t.text or "" for t in p_elem.iter(qn("w:t")))


def direct_run_text_of(p_elem) -> str:
    """Concatenated text of the paragraph's DIRECT ``<w:r>`` children only.

    This is the coordinate space used by the citation detector and the surgical
    inserter, so offsets always match. Runs inside ``<w:hyperlink>`` or inline
    ``<w:sdt>`` are deliberately excluded (modifying them is not safe)."""
    return "".join(
        t.text or ""
        for r in p_elem.findall(qn("w:r"))
        for t in r.findall(qn("w:t")))


# --------------------------------------------------------------------------
# Heading detection
# --------------------------------------------------------------------------
def is_style_heading(paragraph) -> bool:
    """Strong signal: the paragraph is a *styled* heading (Heading 1..9, Title,
    or carries an explicit outline level). Used where a false positive would be
    costly (terminating the reference list, detecting the reference heading)."""
    try:
        style = paragraph.style
        if style is not None and style.name:
            name = style.name.lower()
            if name.startswith("heading") or name == "title":
                return True
    except Exception:
        pass
    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is not None and pPr.find(qn("w:outlineLvl")) is not None:
        return True
    return False


def is_heading(paragraph) -> bool:
    """Heuristic used for *reporting* section names: styled heading first, then
    the bold + short fallback for documents that only use direct formatting."""
    if is_style_heading(paragraph):
        return True
    text = (paragraph.text or "").strip()
    if not text or len(text) >= 90:
        return False
    runs = [r for r in paragraph.runs if (r.text or "").strip()]
    if runs and all(r.bold for r in runs):
        return True
    return False


# --------------------------------------------------------------------------
# Existing-footnote snapshot (input-side, pre-edit)
# --------------------------------------------------------------------------
@dataclass
class FootnoteSnapshot:
    """Snapshot of the footnotes that already exist in the INPUT document."""
    definitions: dict       # {footnote_id: concatenated footnote text}
    body_reference_ids: list  # ids referenced from the body (document order)


def snapshot_existing_footnotes(doc) -> FootnoteSnapshot:
    """Read footnotes.xml (if any) and the body's footnoteReference ids from a
    Document instance. Used to prove existing footnotes survive processing."""
    defs = {}
    part = None
    for p in doc.part.package.iter_parts():
        if str(p.partname).endswith("footnotes.xml"):
            part = p
            break
    if part is not None:
        from lxml import etree
        root = etree.fromstring(part.blob)
        for f in root.findall(qn("w:footnote")):
            fid = f.get(qn("w:id"))
            if fid in ("-1", "0"):
                continue  # separator / continuationSeparator
            text = "".join(t.text or "" for t in f.iter(qn("w:t")))
            defs[fid] = text
    refs = [r.get(qn("w:id")) for r in doc.element.body.iter(qn("w:footnoteReference"))]
    return FootnoteSnapshot(definitions=defs, body_reference_ids=refs)


# --------------------------------------------------------------------------
# DocumentReader
# --------------------------------------------------------------------------
class DocumentReader:
    """Read-only view over one .docx. All consumers share this ordering."""

    def __init__(self, path):
        self.path = str(path)
        self.document = Document(self.path)
        self._paragraphs: Optional[List[Paragraph]] = None
        self._locations: Optional[List[str]] = None

    def all_paragraphs(self) -> List[Paragraph]:
        if self._paragraphs is None:
            info = list(iter_paragraph_info(self.document))
            self._paragraphs = [p for p, _w in info]
            self._locations = [w or "body" for _p, w in info]
        return self._paragraphs

    def locations(self) -> List[str]:
        self.all_paragraphs()
        return self._locations

    def raw_texts(self) -> List[str]:
        """raw_text_of() for every paragraph, in all_paragraphs() order."""
        return [raw_text_of(p._p) for p in self.all_paragraphs()]

    def count_existing_footnotes(self) -> int:
        """Number of existing (non-separator) footnote definitions."""
        return len(snapshot_existing_footnotes(self.document).definitions)

    def existing_footnotes(self) -> FootnoteSnapshot:
        return snapshot_existing_footnotes(self.document)
