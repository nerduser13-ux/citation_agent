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
_SDT_SKIP = (qn("w:sdtPr"), qn("w:alias"))


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
    sdt_where = f"{where} > sdt" if where else "sdt"
    for child in sdt:
        if child.tag in _SDT_SKIP:
            continue
        if child.tag == _P:
            yield Paragraph(child, doc), sdt_where
        elif child.tag == _TBL:
            yield from _iter_table(child, doc, sdt_where)
        elif child.tag == _SDT:
            yield from _iter_sdt(child, doc, sdt_where)
        # w:sdtPr / w:alias and inline-only content are skipped.


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
# Word automatic list numbering (w:numPr / numbering.xml)
# --------------------------------------------------------------------------
def load_numbering_map(doc) -> dict:
    """Parse word/numbering.xml into {numId: {ilvl: (start, numFmt)}}.

    Empty dict when the document has no numbering part."""
    from lxml import etree
    part = None
    for p in doc.part.package.iter_parts():
        if str(p.partname).endswith("numbering.xml"):
            part = p
            break
    if part is None:
        return {}
    root = etree.fromstring(part.blob)
    abstract = {}
    for an in root.findall(qn("w:abstractNum")):
        aid = an.get(qn("w:abstractNumId"))
        lvls = {}
        for lvl in an.findall(qn("w:lvl")):
            ilvl = lvl.get(qn("w:ilvl")) or "0"
            start_el = lvl.find(qn("w:start"))
            fmt_el = lvl.find(qn("w:numFmt"))
            try:
                start = int(start_el.get(qn("w:val"))) if start_el is not None else 1
            except (TypeError, ValueError):
                start = 1
            fmt = fmt_el.get(qn("w:val")) if fmt_el is not None else "decimal"
            lvls[ilvl] = (start, fmt)
        abstract[aid] = lvls
    out = {}
    for num in root.findall(qn("w:num")):
        num_id = num.get(qn("w:numId"))
        aid_el = num.find(qn("w:abstractNumId"))
        if num_id is None or aid_el is None:
            continue
        aid = aid_el.get(qn("w:val"))
        if aid in abstract:
            out[num_id] = abstract[aid]
    return out


def _style_chain_numpr(doc, style_id, _seen=None):
    """numPr element from the style definition (following basedOn), or None."""
    if _seen is None:
        _seen = set()
    if not style_id or style_id in _seen:
        return None
    _seen.add(style_id)
    style = None
    for s in doc.styles.element.findall(qn("w:style")):
        if s.get(qn("w:styleId")) == style_id:
            style = s
            break
    if style is None:
        return None
    pPr = style.find(qn("w:pPr"))
    if pPr is not None:
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            return numPr
    based = style.find(qn("w:basedOn"))
    if based is not None:
        return _style_chain_numpr(doc, based.get(qn("w:val")), _seen)
    return None


def paragraph_numpr_info(p_elem, doc, numbering_map):
    """Visible-number info for a paragraph, or None.

    Returns (numId, ilvl, start_value). The numbering is taken from the
    paragraph's own <w:numPr> or, failing that, from its paragraph style
    chain (this is how "List Number" styled paragraphs are numbered). Only
    decimal numbering qualifies (bullets/letters/romans are not reference
    numbers)."""
    if not numbering_map:
        return None
    numPr = None
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is not None:
        numPr = pPr.find(qn("w:numPr"))
        if numPr is None:
            pstyle = pPr.find(qn("w:pStyle"))
            if pstyle is not None:
                numPr = _style_chain_numpr(doc, pstyle.get(qn("w:val")))
    if numPr is None:
        return None
    numId_el = numPr.find(qn("w:numId"))
    ilvl_el = numPr.find(qn("w:ilvl"))
    if numId_el is None:
        return None
    num_id = numId_el.get(qn("w:val"))
    if num_id in (None, "0"):
        return None
    ilvl = (ilvl_el.get(qn("w:val")) if ilvl_el is not None else "0") or "0"
    lvls = numbering_map.get(num_id)
    if not lvls:
        return None
    start, fmt = lvls.get(ilvl, (1, "decimal"))
    if fmt != "decimal":
        return None
    return num_id, ilvl, start


def compute_list_numbers(paragraphs, doc, numbering_map) -> dict:
    """Map paragraph index -> visible automatic number (document-wide
    sequence per list, as rendered by Word). Only decimal lists."""
    seq = {}
    out = {}
    for i, p in enumerate(paragraphs):
        info = paragraph_numpr_info(p._p, doc, numbering_map)
        if info is None:
            continue
        num_id, ilvl, start = info
        key = (num_id, ilvl)
        n = seq.get(key, start)
        seq[key] = n + 1
        out[i] = n
    return out


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

    def list_numbers(self) -> dict:
        """{paragraph_index: visible automatic number} for paragraphs that
        belong to a Word decimal numbered list (direct or via their style).
        Lets the reference parser recognize lists whose numbers are rendered
        by Word's numbering engine instead of being typed into the text."""
        self.all_paragraphs()
        return compute_list_numbers(
            self._paragraphs, self.document, load_numbering_map(self.document))

    def count_existing_footnotes(self) -> int:
        """Number of existing (non-separator) footnote definitions."""
        return len(snapshot_existing_footnotes(self.document).definitions)

    def existing_footnotes(self) -> FootnoteSnapshot:
        return snapshot_existing_footnotes(self.document)

