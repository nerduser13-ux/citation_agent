"""word_footnotes.py

Inserts REAL Microsoft Word footnotes while preserving the existing document as
much as technically possible.

Two backends:
  * XmlFootnoteBackend  - cross-platform OOXML manipulation (DEFAULT, validated).
  * process_document_com - Windows Microsoft Word COM automation (EXPERIMENTAL;
                           requires Windows + Microsoft Word + pywin32; NOT
                           executed in the Linux build/test environment).

Surgical-editing principles:
  * We never rebuild a whole paragraph. Only the single ``<w:r>`` run that
    actually contains a citation marker is replaced - and only when that run
    is "simple" (its children are exclusively ``w:rPr``/``w:t``, plus Word's
    content-free ``w:lastRenderedPageBreak`` hint, which is re-emitted at its
    exact character position). Every other
    paragraph child (bookmarks, hyperlinks, fields, drawings, proof errors,
    content controls) is left exactly where it is, in the same order.
  * Splitting happens at the text level: the marker's before/after pieces are
    new runs carrying a deep copy of the original run properties, so bold,
    italic, underline, fonts, sizes, colours and character styles survive.
  * The footnote-reference run keeps the original run's character formatting,
    gains the FootnoteReference character style, and is explicitly marked
    superscript so Word renders the footnote reference correctly.
  * Several references cited together ([1,2], [1-3], [1]-[3]) become footnote
    references side by side, with a separator between them (a space by
    default, formatted like the references) so that 1 and 2 never read as 12.
    Citations separated only by spaces and a comma or semicolon ("[1], [2]",
    "[1][2]") are joined the same way in replace mode: the separator text
    between them is removed - also across Word's invisible run splits, but only
    through plain text runs, never across words, links, fields or footnotes.
  * Word often stores one marker in several runs (editing history, a
    spell-check mark, its hidden _GoBack bookmark, a language tag): "[3]" in
    one run and "-[4]" in the next. Such a marker is processed when every
    piece is plain, visible text and nothing but spell-check marks or
    bookmarks sits between the pieces; its text is removed from each piece
    and each piece keeps its own formatting for the text around the marker.
  * A marker interrupted by a tab, picture, link or field, partly hidden or
    struck through, or inside a hyperlink / a complex-field result
    (reference-manager citation) / a run with a drawing, field, object or
    other non-text content, is NOT modified - it is reported AMBIGUOUS
    instead of risking corruption.
  * Idempotency: a marker immediately flanked by a footnote reference whose
    content equals one of the citation's own URLs is treated as previously
    processed (SKIPPED); unrelated pre-existing footnotes never cause a skip.
"""

import copy
import os
from collections import defaultdict

from citation_detector import JOIN_GAP_RE

from lxml import etree

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part


FOOTNOTES_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
)


FOOTNOTES_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:footnotes '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:footnote w:type="separator" w:id="-1">'
    '<w:p><w:r><w:separator/></w:r></w:p>'
    '</w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0">'
    '<w:p><w:r><w:continuationSeparator/></w:r></w:p>'
    '</w:footnote>'
    "</w:footnotes>"
)


# --------------------------------------------------------------------------
# Styles
# --------------------------------------------------------------------------

def ensure_footnote_styles(doc):
    """Make sure the FootnoteReference and FootnoteText styles exist.

    Reuses an existing style of the same name if present; otherwise creates
    the standard styles.

    Returns:
        (reference_style_id, footnote_text_style_id)
    """
    styles_root = doc.styles.element

    def find_by_name(name):
        for st in styles_root.findall(qn("w:style")):
            nm = st.find(qn("w:name"))
            if nm is not None and (
                nm.get(qn("w:val")) or ""
            ).lower() == name:
                sid = st.get(qn("w:styleId"))
                if sid:
                    return sid
        return None

    ref_id = find_by_name("footnote reference")
    text_id = find_by_name("footnote text")

    def add(style_id, style_type, name_val):
        st = OxmlElement("w:style")
        st.set(qn("w:type"), style_type)
        st.set(qn("w:styleId"), style_id)

        nm = OxmlElement("w:name")
        nm.set(qn("w:val"), name_val)

        st.append(nm)
        styles_root.append(st)

    if ref_id is None:
        ref_id = "FootnoteReference"
        add(ref_id, "character", "footnote reference")

    if text_id is None:
        text_id = "FootnoteText"
        add(text_id, "paragraph", "footnote text")

    return ref_id, text_id


# --------------------------------------------------------------------------
# Low-level OOXML helpers
# --------------------------------------------------------------------------

def _set_rstyle(rpr, style_id):
    """Return rPr with w:rStyle set to style_id.

    rStyle is inserted as the first child of rPr, as required by the OOXML
    schema.
    """
    if rpr is None:
        rpr = OxmlElement("w:rPr")

    existing = rpr.find(qn("w:rStyle"))

    if existing is not None:
        existing.set(qn("w:val"), style_id)
        rpr.remove(existing)
    else:
        existing = OxmlElement("w:rStyle")
        existing.set(qn("w:val"), style_id)

    rpr.insert(0, existing)

    return rpr


# CT_RPr children that the OOXML schema places AFTER w:vertAlign. w:rPr is a
# strict sequence; appending vertAlign after e.g. w:lang (present in most
# Word-authored runs) produces a schema-invalid run property list.
_AFTER_VERTALIGN = tuple(
    qn(f"w:{name}")
    for name in (
        "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish",
        "oMath", "rPrChange",
    )
)


def _set_superscript(rpr):
    """Add/replace w:vertAlign so the run is explicitly superscript.

    This is important for the footnote reference marker. Word normally
    renders a footnoteReference as superscript, but explicitly writing
    w:vertAlign makes the OOXML unambiguous and avoids relying on the
    document's style definitions. A new element is inserted at its
    schema-correct position within w:rPr.
    """
    if rpr is None:
        rpr = OxmlElement("w:rPr")

    vert_align = rpr.find(qn("w:vertAlign"))

    if vert_align is None:
        vert_align = OxmlElement("w:vertAlign")
        follower = next(
            (child for child in rpr if child.tag in _AFTER_VERTALIGN),
            None,
        )
        if follower is not None:
            follower.addprevious(vert_align)
        else:
            rpr.append(vert_align)

    vert_align.set(qn("w:val"), "superscript")

    return rpr


def _make_t(text):
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    return t


def make_text_run(text, rpr, anchors=()):
    """Create a plain text run.

    The original rPr is deep-copied so modifying the new run never modifies
    the original source run.

    ``anchors`` is an optional list of ``(offset, element)`` pairs (offsets
    local to ``text``, in document order): zero-width run children such as
    ``w:lastRenderedPageBreak`` that are re-emitted exactly before the
    character they originally preceded.
    """
    r = OxmlElement("w:r")

    if rpr is not None:
        r.append(copy.deepcopy(rpr))

    prev = 0

    for off, el in anchors:
        if off > prev:
            r.append(_make_t(text[prev:off]))
            prev = off
        r.append(copy.deepcopy(el))

    if prev < len(text):
        r.append(_make_t(text[prev:]))

    return r


def make_anchor_run(anchor, rpr):
    """A run holding only a zero-width element whose text position was
    removed (it sat inside a replaced marker) - so it is never dropped."""
    r = OxmlElement("w:r")

    if rpr is not None:
        r.append(copy.deepcopy(rpr))

    r.append(copy.deepcopy(anchor))

    return r


def make_ref_run(fid, rpr, ref_style_id):
    """Create the genuine Word footnote-reference run.

    The generated XML is conceptually:

        <w:r>
          <w:rPr>
            ...original formatting...
            <w:rStyle w:val="FootnoteReference"/>
            <w:vertAlign w:val="superscript"/>
          </w:rPr>
          <w:footnoteReference w:id="..."/>
        </w:r>

    The original character formatting is preserved, while the footnote
    reference is explicitly superscripted.
    """
    run = OxmlElement("w:r")

    # Start with the original character formatting.
    new_rpr = copy.deepcopy(rpr) if rpr is not None else None

    # Add the FootnoteReference character style.
    new_rpr = _set_rstyle(new_rpr, ref_style_id)

    # Explicitly make the footnote reference superscript.
    new_rpr = _set_superscript(new_rpr)

    run.append(new_rpr)

    # Genuine Word footnote reference.
    footnote_ref = OxmlElement("w:footnoteReference")
    footnote_ref.set(qn("w:id"), str(fid))

    run.append(footnote_ref)

    return run


def make_separator_run(text, rpr, ref_style_id):
    """Text placed between two footnote references (e.g. a space or comma),
    formatted exactly like the references (superscript) so they read as one
    group: 1 2 (or 1,2) instead of 12."""
    new_rpr = copy.deepcopy(rpr) if rpr is not None else None
    new_rpr = _set_superscript(_set_rstyle(new_rpr, ref_style_id))
    run = OxmlElement("w:r")
    run.append(new_rpr)
    run.append(_make_t(text))
    return run


def add_footnote_def(root, fid, url, ref_style_id, text_style_id):
    """Append a genuine <w:footnote> definition containing exactly ``url``.

    The paragraph carries the FootnoteText paragraph style.

    The first run contains the automatic Word footnote number.

    The second run contains the exact URL.
    """
    ft = OxmlElement("w:footnote")
    ft.set(qn("w:id"), str(fid))

    p = OxmlElement("w:p")

    # Paragraph properties.
    pPr = OxmlElement("w:pPr")

    ps = OxmlElement("w:pStyle")
    ps.set(qn("w:val"), text_style_id)

    pPr.append(ps)
    p.append(pPr)

    # ---------------------------------------------------------------
    # Automatic footnote number inside footnotes.xml
    # ---------------------------------------------------------------
    r1 = OxmlElement("w:r")

    r1_rpr = OxmlElement("w:rPr")

    # FootnoteReference character style.
    rstyle = OxmlElement("w:rStyle")
    rstyle.set(qn("w:val"), ref_style_id)
    r1_rpr.append(rstyle)

    # Explicit superscript formatting.
    vert_align = OxmlElement("w:vertAlign")
    vert_align.set(qn("w:val"), "superscript")
    r1_rpr.append(vert_align)

    r1.append(r1_rpr)

    footnote_ref = OxmlElement("w:footnoteRef")
    r1.append(footnote_ref)

    p.append(r1)

    # ---------------------------------------------------------------
    # Exact URL text
    # ---------------------------------------------------------------
    r2 = OxmlElement("w:r")

    t2 = OxmlElement("w:t")
    t2.set(qn("xml:space"), "preserve")
    t2.text = url

    r2.append(t2)
    p.append(r2)

    ft.append(p)
    root.append(ft)


# --------------------------------------------------------------------------
# OOXML backend
# --------------------------------------------------------------------------

class XmlFootnoteBackend:
    """Cross-platform OOXML footnote backend."""

    def __init__(self, doc, style_ids=None):
        self.doc = doc
        self.package = doc.part.package

        self.style_ids = style_ids or (
            "FootnoteReference",
            "FootnoteText",
        )

        self.part, self.root = self._ensure_part()

        self.next_id = self._compute_next_id(self.root)

        self.inserted = 0

        # {fid: url}
        self.new_footnotes = {}

        self._footnote_text = self._build_text_map()

    def _ensure_part(self):
        """Find existing footnotes.xml or create one."""
        for part in self.package.iter_parts():
            if str(part.partname).endswith("footnotes.xml"):
                return part, etree.fromstring(part.blob)

        part = Part(
            PackURI("/word/footnotes.xml"),
            FOOTNOTES_CT,
            FOOTNOTES_TEMPLATE.encode("utf-8"),
            self.package,
        )

        self.doc.part.relate_to(part, RT.FOOTNOTES)

        return part, etree.fromstring(
            FOOTNOTES_TEMPLATE.encode("utf-8")
        )

    @staticmethod
    def _compute_next_id(root):
        """Find the next available positive footnote ID."""
        ids = []

        for f in root.findall(qn("w:footnote")):
            try:
                ids.append(int(f.get(qn("w:id"))))
            except (TypeError, ValueError):
                pass

        return max(ids) + 1 if ids else 1

    def _build_text_map(self):
        """Build {footnote_id: footnote_text}."""
        mapping = {}

        for f in self.root.findall(qn("w:footnote")):
            fid = f.get(qn("w:id"))

            if fid in (None, "-1", "0"):
                continue

            mapping[fid] = "".join(
                t.text or ""
                for t in f.iter(qn("w:t"))
            )

        return mapping

    def footnote_text(self, fid):
        """Return text of an existing/new footnote."""
        return self._footnote_text.get(str(fid))

    def add_footnote(self, url):
        """Create a new footnote and return its numeric ID."""
        fid = self.next_id
        self.next_id += 1

        add_footnote_def(
            self.root,
            fid,
            url,
            *self.style_ids,
        )

        self._footnote_text[str(fid)] = url
        self.new_footnotes[fid] = url
        self.inserted += 1

        return fid

    def save(self):
        """Write the in-memory footnotes.xml back into the package."""
        self.part._blob = etree.tostring(
            self.root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )


# --------------------------------------------------------------------------
# Safety helpers
# --------------------------------------------------------------------------

# Zero-width run children that carry no content and may travel with a split.
# w:lastRenderedPageBreak is Word's layout-cache hint ("a page break fell here
# when the file was last saved"); Word writes it into practically every
# multi-page document, so treating it as unsafe content would leave a large
# share of real citations unprocessed. It is preserved at its exact position.
_ZERO_WIDTH_RUN_CHILDREN = (qn("w:lastRenderedPageBreak"),)

_SIMPLE_RUN_CHILDREN = (qn("w:rPr"), qn("w:t")) + _ZERO_WIDTH_RUN_CHILDREN


def _run_is_simple(run_elem):
    """Return True only for runs containing rPr/t children (plus zero-width
    layout hints, see ``_ZERO_WIDTH_RUN_CHILDREN``).

    Runs containing drawings, fields, breaks, tabs, objects, etc. are not safe
    to split.
    """
    return all(
        child.tag in _SIMPLE_RUN_CHILDREN
        for child in run_elem
    )


def _zero_width_anchors(run_elem):
    """``[(text_offset, element)]`` for the run's zero-width children, in
    document order (offsets are measured over the run's w:t text)."""
    anchors = []
    pos = 0

    for child in run_elem:
        if child.tag == qn("w:t"):
            pos += len(child.text or "")
        elif child.tag in _ZERO_WIDTH_RUN_CHILDREN:
            anchors.append((pos, child))

    return anchors


# Textless elements that may sit between two citations that are joined.
_JOIN_TRANSPARENT = (qn("w:proofErr"), qn("w:bookmarkStart"), qn("w:bookmarkEnd"))


def _only_text_between(first, last):
    """True if every sibling strictly between run elements ``first`` and
    ``last`` is a plain text run or a spell-check mark / bookmark."""
    node = first
    while node is not last:
        node = node.getnext()
        if node is None:
            return False
        if node is last:
            break
        if node.tag == qn("w:r"):
            if not _run_is_simple(node):
                return False
        elif node.tag not in _JOIN_TRANSPARENT:
            return False
    return True


_ON = (None, "1", "true", "on")


def _hides_or_strikes(run_elem):
    """Hidden (w:vanish) or struck-through text changes what the reader sees,
    so a marker partly made of it is not processed."""
    rpr = run_elem.find(qn("w:rPr"))
    if rpr is None:
        return False
    return any(
        el is not None and el.get(qn("w:val")) in _ON
        for el in (rpr.find(qn("w:vanish")), rpr.find(qn("w:strike")),
                   rpr.find(qn("w:dstrike")))
    )


def _joinable(prev, cit, full):
    """May citation ``cit`` be placed side by side with ``prev`` (removing the
    text between them)? Only if that text is spaces plus at most one comma or
    semicolon, and everything between the two host runs is plain text runs
    (or spell-check marks / bookmarks) - never a link, field, tab, footnote..."""
    if not JOIN_GAP_RE.fullmatch(full[prev.end:cit.start]):
        return False
    return _only_text_between(prev.host, cit.host)


def _is_ref_run(run_elem):
    """Return True when the run contains a footnoteReference."""
    return (
        run_elem.tag == qn("w:r")
        and run_elem.find(qn("w:footnoteReference")) is not None
    )


# --------------------------------------------------------------------------
# Paragraph processing
# --------------------------------------------------------------------------

def process_paragraph(
    p_elem,
    para_citations,
    backend,
    replace_marker,
    ref_map,
    separator=" ",
):
    """Insert footnotes for citations belonging to one paragraph.

    Only the individual run containing a citation is replaced (plus, when
    citations are joined, the plain text runs holding the ", " between them).
    Everything else in the paragraph remains untouched. ``separator`` goes
    between footnote references that stand side by side ("" = none).
    """

    runs = p_elem.findall(qn("w:r"))

    run_info = []
    full = ""

    for r in runs:
        txt = "".join(
            t.text or ""
            for t in r.findall(qn("w:t"))
        )

        run_info.append({
            "elem": r,
            "text": txt,
            "start": len(full),
            "end": len(full) + len(txt),
            "rpr": r.find(qn("w:rPr")),
        })

        full += txt

    # ------------------------------------------------------------------
    # 1. Map every citation to a safe host run.
    # ------------------------------------------------------------------

    for c in para_citations:
        if c.status != "pending":
            continue

        # Every run the marker touches (Word may store "[3]-[4]" as
        # "[3]" + "-[4]"), including runs without text inside it.
        covering = [
            ri for ri in run_info
            if ri["start"] < c.end and c.start < ri["end"]
        ]

        if not covering:
            c.status = "AMBIGUOUS"
            c.note = "Citation not found in the paragraph's runs; left unchanged"
            continue

        host = covering[-1]          # the footnote references go here

        if len(covering) == 1:
            if not _run_is_simple(host["elem"]):
                c.status = "AMBIGUOUS"
                c.note = (
                    "Citation run contains non-text content "
                    "(field/drawing/break/...); left unchanged"
                )
                continue
        elif not all(_run_is_simple(ri["elem"]) for ri in covering):
            c.status = "AMBIGUOUS"
            c.note = (
                "Citation is split by a tab, line break, picture or other "
                "non-text element; left unchanged"
            )
            continue
        elif any(_hides_or_strikes(ri["elem"]) for ri in covering):
            c.status = "AMBIGUOUS"
            c.note = ("Citation is partly hidden or struck through; "
                      "left unchanged")
            continue
        elif not _only_text_between(covering[0]["elem"], host["elem"]):
            c.status = "AMBIGUOUS"
            c.note = ("Citation is interrupted by a link, field or similar; "
                      "left unchanged")
            continue

        # --------------------------------------------------------------
        # Idempotency check.
        # --------------------------------------------------------------

        for sib in (
            covering[0]["elem"].getprevious(),
            host["elem"].getnext(),
        ):
            if sib is not None and _is_ref_run(sib):
                footnote_ref = sib.find(
                    qn("w:footnoteReference")
                )

                if footnote_ref is None:
                    continue

                fid = footnote_ref.get(qn("w:id"))
                text = backend.footnote_text(fid)

                if any(
                    text == ref_map[n]["url"]
                    for n in c.numbers
                    if n in ref_map
                    and ref_map[n]["url"]
                ):
                    c.status = "SKIPPED"
                    c.note = (
                        "Previously processed "
                        "(matching footnote already present)"
                    )
                    break

        if c.status != "pending":
            continue

        c.host = host["elem"]

        c.local_start = max(c.start, host["start"]) - host["start"]
        c.local_end = c.end - host["start"]

        # The marker's text in earlier runs: removed in replace mode.
        c.pieces = [
            (ri["elem"], c.start - ri["start"] if c.start > ri["start"] else 0,
             ri["end"] - ri["start"])
            for ri in covering[:-1] if ri["end"] > ri["start"]
        ]

    # ------------------------------------------------------------------
    # 2. Validate citation numbers and allocate footnote IDs.
    # ------------------------------------------------------------------

    for c in para_citations:
        if c.status != "pending":
            continue

        c.fids = []

        ok = True

        for num in c.numbers:
            if num not in ref_map:
                ok = False
                c.note = (
                    f"Reference {num} not found in reference section"
                )

            elif not ref_map[num]["url"]:
                ok = False
                c.note = (
                    f"Reference {num} has no URL in its entry"
                )

        if not ok:
            c.status = "ERROR"
            continue

        c.fids = [
            (
                num,
                backend.add_footnote(
                    ref_map[num]["url"]
                ),
            )
            for num in c.numbers
        ]

    # ------------------------------------------------------------------
    # 3. Join neighbours: "[1], [2]" -> footnotes 1 and 2 side by side.
    #    Replace mode only (in keep-marker mode no text is removed). The
    #    text between them is cut from every run it lies in.
    # ------------------------------------------------------------------

    cuts = defaultdict(list)            # run -> [(local_lo, local_hi)]

    if replace_marker:
        for c in para_citations:        # pieces of markers split over runs
            if c.status == "pending" and c.host is not None and c.fids:
                for elem, lo, hi in c.pieces:
                    cuts[elem].append((lo, hi))

        prev = None
        for c in sorted((c for c in para_citations if c.start >= 0),
                        key=lambda c: c.start):
            ready = c.status == "pending" and c.host is not None and bool(c.fids)
            if ready and prev is not None and _joinable(prev, c, full):
                c.joined = True
                c.join_from = prev.end
                for ri in run_info:
                    lo = max(prev.end, ri["start"])
                    hi = min(c.start, ri["end"])
                    if lo < hi:
                        cuts[ri["elem"]].append(
                            (lo - ri["start"], hi - ri["start"]))
            prev = c if ready else None

    # ------------------------------------------------------------------
    # 4. Replace the citation-containing runs (and cut the joining text).
    # ------------------------------------------------------------------

    groups = defaultdict(list)

    for c in para_citations:
        if (
            c.host is not None
            and c.status == "pending"
            and c.fids
        ):
            groups[c.host].append(c)

    targets = list(groups) + [r for r in cuts if r not in groups]

    for run_elem in targets:

        cites = sorted(
            groups.get(run_elem, []),
            key=lambda c: c.local_start
        )

        new_elems = _build_run_replacement(
            run_elem,
            cites,
            replace_marker,
            backend.style_ids[0],
            separator,
            sorted(cuts.get(run_elem, [])),
        )

        parent = run_elem.getparent()

        idx = list(parent).index(run_elem)

        parent.remove(run_elem)

        for j, el in enumerate(new_elems):
            parent.insert(
                idx + j,
                el,
            )

        for c in cites:
            c.status = "SUCCESS"


def _build_run_replacement(
    run_elem,
    cites,
    replace_marker,
    ref_style_id,
    separator=" ",
    cuts=(),
):
    """Build replacement elements for one citation-containing run.

    Replace mode:

        [before] [footnote ref] [sep] [footnote ref] [after]

    Keep-marker mode:

        [before] [marker] [footnote ref] [sep] [footnote ref] [after]

    ``cuts`` are local text ranges to remove: the ", " between two joined
    citations. A joined citation's references start with a separator, so
    "[1], [2]" becomes [ref 1] [sep] [ref 2].

    Text runs preserve the original rPr.

    Footnote reference runs preserve the original formatting but explicitly
    receive the FootnoteReference style and superscript formatting.

    Zero-width children (w:lastRenderedPageBreak) are re-emitted before the
    same character they preceded in the original run. One that sat inside a
    replaced marker is kept as its own run where the marker was, and one at
    the very end of the run stays at the end - none is ever dropped.
    """

    text = "".join(
        t.text or ""
        for t in run_elem.findall(qn("w:t"))
    )

    rpr = run_elem.find(qn("w:rPr"))

    anchors = _zero_width_anchors(run_elem)
    placed = set()

    def take_anchors(lo, hi, include_hi=False):
        """Claim not-yet-placed anchors with lo <= offset < hi (or <= hi)."""
        out = []
        for i, (off, el) in enumerate(anchors):
            if i in placed:
                continue
            if lo <= off < hi or (include_hi and off == hi):
                placed.add(i)
                out.append((off, el))
        return out

    def text_run(lo, hi, include_hi=False):
        return make_text_run(
            text[lo:hi],
            rpr,
            [(off - lo, el) for off, el in take_anchors(lo, hi, include_hi)],
        )

    def removed(lo, hi):
        """Zero-width elements of removed text stay at the removal point."""
        return [make_anchor_run(el, rpr) for _off, el in take_anchors(lo, hi)]

    # Citations and cuts in text order (a cut comes before the marker it
    # leads up to; the two never overlap).
    ops = sorted(
        [(lo, hi, None) for lo, hi in cuts]
        + [(c.local_start, c.local_end, c) for c in cites],
        key=lambda op: (op[0], op[2] is not None),
    )

    new_elems = []

    pos = 0

    for ls, le, c in ops:

        # --------------------------------------------------------------
        # Text before the citation / cut.
        # --------------------------------------------------------------

        if ls > pos:
            new_elems.append(text_run(pos, ls))

        if c is None:
            # The ", " between two joined citations disappears.
            new_elems.extend(removed(ls, le))
            pos = le
            continue

        # --------------------------------------------------------------
        # Keep [n] marker if requested (otherwise it disappears, and any
        # zero-width element that sat inside it stays at this position).
        # --------------------------------------------------------------

        if not replace_marker:
            if le > ls:
                new_elems.append(text_run(ls, le))
        else:
            new_elems.extend(removed(ls, le))

        # --------------------------------------------------------------
        # Genuine Word footnote references, side by side with separators.
        # --------------------------------------------------------------

        for k, (_num, fid) in enumerate(c.fids):
            if separator and (k > 0 or c.joined):
                new_elems.append(
                    make_separator_run(separator, rpr, ref_style_id)
                )
            new_elems.append(
                make_ref_run(
                    fid,
                    rpr,
                    ref_style_id,
                )
            )

        pos = le

    # ------------------------------------------------------------------
    # Text after the last citation.
    # ------------------------------------------------------------------

    if pos < len(text):
        new_elems.append(text_run(pos, len(text), include_hi=True))

    # Anything still unplaced sat at the very end of a run that ended with
    # a citation marker: keep it at the end.
    new_elems.extend(
        make_anchor_run(el, rpr)
        for _off, el in take_anchors(0, len(text), include_hi=True)
    )

    return new_elems


# --------------------------------------------------------------------------
# Windows Microsoft Word COM backend
# --------------------------------------------------------------------------

def process_document_com(
    path,
    citations,
    ref_map,
    replace_marker,
    out_path,
    separator=" ",
):
    """Use Microsoft Word COM automation to insert footnotes.

    Requires:
      - Windows
      - Microsoft Word
      - pywin32

    EXPERIMENTAL - not executable in the Linux build/test environment. Written
    against the documented Word object model:
      * ``Range.Find.Execute()`` redefines the *Range itself* to the match
        (the Find object has no Range property);
      * ``Footnotes.Add(Range, Reference, Text)`` - the 2nd positional
        argument is a *custom reference mark*, so the URL is passed as
        ``Text=`` to get an automatically numbered footnote.

    Markers are matched by text search in document order. Every positioned
    citation - including ones classified AMBIGUOUS / SKIPPED - is searched
    for, so the search cursor stays aligned with the detector's sequence and a
    later identical marker is never matched to an earlier, unsafe occurrence.
    Citations inside hyperlinks (no position) are not searched for.

    Footnotes of one citation ([1,2], [1-3]) are separated by ``separator``
    (made superscript); in replace mode a citation that follows the previous
    one with only spaces and a comma/semicolon between ("[1], [2]") is joined
    to it, like the XML backend does.
    """

    if os.name != "nt":
        raise RuntimeError(
            "--backend com requires Windows with Microsoft Word installed "
            "(pywin32). Use the default XML backend on other platforms."
        )

    import win32com.client

    word = win32com.client.Dispatch(
        "Word.Application"
    )

    try:
        word.Visible = False

        doc = word.Documents.Open(
            os.path.abspath(path)
        )

        try:
            last_pos = 0
            prev_done = None        # previous positioned citation, if inserted here

            positioned = sorted(
                (c for c in citations if c.start >= 0),
                key=lambda c: (c.para_index, c.start),
            )

            for c in positioned:

                rng = doc.Range(
                    last_pos,
                    doc.Content.End,
                )

                fnd = rng.Find

                # Find options are sticky for the whole Word session (e.g. a
                # user's earlier wildcard search would turn "[1]" into a
                # character class) - set every relevant option explicitly.
                fnd.ClearFormatting()
                fnd.Text = c.marker_text
                fnd.Forward = True
                fnd.Wrap = 0                    # wdFindStop
                fnd.Format = False
                fnd.MatchCase = False
                fnd.MatchWholeWord = False
                fnd.MatchWildcards = False
                fnd.MatchSoundsLike = False
                fnd.MatchAllWordForms = False

                if not fnd.Execute():
                    if c.status == "pending":
                        c.status = "AMBIGUOUS"
                        c.note = (
                            "Marker not found by Word COM; "
                            "left unchanged"
                        )
                    prev_done = None
                    continue

                # ``rng`` now spans exactly the matched marker text.
                if c.status != "pending":
                    last_pos = rng.End      # unsafe marker: step over it
                    prev_done = None
                    continue

                # Validate every number BEFORE touching the document so a
                # range with one bad number never gets partial footnotes.
                bad = next(
                    (n for n in c.numbers
                     if not ref_map.get(n, {}).get("url")),
                    None,
                )

                if bad is not None:
                    c.status = "ERROR"
                    c.note = (
                        f"Reference {bad} "
                        + (
                            "not found in reference section"
                            if bad not in ref_map
                            else "has no URL in its entry"
                        )
                    )
                    last_pos = rng.End
                    prev_done = None
                    continue

                # "[1], [2]": join to the previous citation's footnotes when
                # only spaces and a comma/semicolon separate them.
                joined = (
                    replace_marker
                    and prev_done is not None
                    and prev_done.para_index == c.para_index
                    and bool(JOIN_GAP_RE.fullmatch(
                        doc.Range(last_pos, rng.Start).Text or ""))
                )

                # Remove the marker first (replace mode), then insert every
                # footnote at a collapsed range, each one AFTER the previous
                # reference mark, so [4]-[6] yields refs 4, 5, 6 in order.
                if replace_marker:
                    pos = last_pos if joined else rng.Start
                    doc.Range(pos, rng.End).Delete()
                else:
                    pos = rng.End

                for k, num in enumerate(c.numbers):
                    if separator and (k > 0 or joined):
                        sep_rng = doc.Range(pos, pos)
                        sep_rng.InsertAfter(separator)
                        sep_rng.Font.Superscript = True
                        pos = sep_rng.End
                    fn = doc.Footnotes.Add(
                        Range=doc.Range(pos, pos),
                        Text=ref_map[num]["url"],
                    )
                    pos = fn.Reference.End

                last_pos = pos

                c.status = "SUCCESS"
                c.note = "COM footnote inserted"
                c.joined = joined
                c.join_from = prev_done.end if joined else -1
                prev_done = c

            doc.SaveAs(
                os.path.abspath(out_path)
            )

        finally:
            doc.Close(
                SaveChanges=False
            )

    finally:
        word.Quit()
