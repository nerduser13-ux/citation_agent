"""citation_detector.py

Finds numbered citation markers in the document BODY only (the reference
section itself is excluded). Supported marker syntax:

    [1]
    [2]
    [1] and [2]
    [1], [2], and [3]
    [1]-[3]      (hyphen)
    [1]–[3]      (en dash)
    [1]—[3]      (em dash)
    [2] ... [2]  (repeated citations -> one Citation object per occurrence)

Ranges expand correctly: [4]-[6] -> 4, 5, 6 (three footnote references).

COORDINATE SYSTEM: offsets are measured over the paragraph's DIRECT ``<w:r>``
run text only (``direct_run_text_of``). This matches the coordinate space of
the surgical inserter in word_footnotes.py, so a marker after a hyperlink is
never mis-attributed to the hyperlink's runs.

SAFETY: markers inside ``<w:hyperlink>`` elements are reported (AMBIGUOUS) but
never modified. Markers that are (part of) the displayed result of a complex
field - how EndNote, Zotero and Mendeley Desktop store "[1]", and how
cross-references work - are reported AMBIGUOUS: a footnote inserted there would
be destroyed (or break the field) the next time the field is refreshed.
Markers that span runs are resolved as AMBIGUOUS at processing time. Reversed
ranges ([3]-[1]) are reported AMBIGUOUS rather than guessed.

Statuses:
    pending    -> not yet processed (resolved by word_footnotes.process_paragraph)
    SUCCESS    -> footnote(s) inserted
    ERROR      -> reference number missing, or reference has no URL
    SKIPPED    -> already processed by a previous run (idempotency guard)
    AMBIGUOUS  -> cannot be safely mapped/modified; left untouched
"""
import re
from dataclasses import dataclass, field

from docx.oxml.ns import qn

from document_reader import direct_run_text_of

STATUS_PENDING = "pending"
STATUS_SUCCESS = "SUCCESS"
STATUS_ERROR = "ERROR"
STATUS_SKIPPED = "SKIPPED"
STATUS_AMBIGUOUS = "AMBIGUOUS"

_W_R = qn("w:r")
_W_T = qn("w:t")
_FLDCHAR = qn("w:fldChar")
_FLDCHAR_TYPE = qn("w:fldCharType")
_TXBX = qn("w:txbxContent")


def _apply_field_chars(elem, depth):
    """Update the complex-field nesting ``depth`` with every w:fldChar under
    ``elem`` (document order). Text boxes are skipped: their fields are
    self-contained and belong to a different story."""
    stack = [elem]
    while stack:
        node = stack.pop()
        if node.tag == _TXBX:
            continue
        if node.tag == _FLDCHAR:
            kind = node.get(_FLDCHAR_TYPE)
            if kind == "begin":
                depth += 1
            elif kind == "end":
                depth = max(0, depth - 1)
            continue
        stack.extend(reversed(list(node)))
    return depth


def field_result_spans(p_elem, depth):
    """Direct-run coordinate spans ``[(start, end)]`` of the paragraph's runs
    that sit inside a complex field (between fldChar begin and end), plus the
    nesting depth after the paragraph (fields may span paragraphs)."""
    spans = []
    pos = 0
    for child in p_elem:
        if child.tag == _W_R:
            length = sum(len(t.text or "") for t in child.findall(_W_T))
            if depth > 0 and length:
                spans.append((pos, pos + length))
            pos += length
        depth = _apply_field_chars(child, depth)
    return spans, depth


@dataclass
class Citation:
    para_index: int          # index into the shared paragraph list
    start: int               # direct-run coordinate offset of the marker
    end: int                 # offset just past the marker
    numbers: list            # reference numbers this marker maps to
    is_range: bool           # True for [a]-[b] style ranges
    marker_text: str         # exact matched text, e.g. "[1]" or "[4]-[6]"
    status: str = STATUS_PENDING
    note: str = ""
    fids: list = field(default_factory=list)    # (number, footnote_id) pairs
    host = None              # run element containing the marker (set later)
    local_start: int = 0
    local_end: int = 0


class CitationDetector:
    TOKEN_RE = re.compile(r"\[\d+\]")
    # Separator between two range tokens: optional whitespace, one dash
    # (hyphen, en dash or em dash), optional whitespace.
    RANGE_GAP_RE = re.compile(r"^\s*[-\u2013\u2014]\s*$")

    def detect(self, paragraphs, ref_section_start):
        """Detect citation markers in body paragraphs only
        (index < ref_section_start)."""
        citations = []
        field_depth = 0   # complex-field nesting, carried across paragraphs
        for idx, p in enumerate(paragraphs):
            if idx >= ref_section_start:
                break
            t = direct_run_text_of(p._p)
            in_field, field_depth = field_result_spans(p._p, field_depth)
            tokens = [
                (m.start(), m.end(), int(m.group()[1:-1]))
                for m in self.TOKEN_RE.finditer(t)
            ]
            para_cites = []
            i = 0
            while i < len(tokens):
                s, e, a = tokens[i]
                if i + 1 < len(tokens):
                    s2, e2, b = tokens[i + 1]
                    if self.RANGE_GAP_RE.match(t[e:s2]):
                        if a <= b:
                            para_cites.append(Citation(
                                idx, s, e2, list(range(a, b + 1)), True, t[s:e2]))
                        else:
                            para_cites.append(Citation(
                                idx, s, e2, [a], True, t[s:e2],
                                status=STATUS_AMBIGUOUS,
                                note="Reversed range (start > end); left unchanged"))
                        i += 2
                        continue
                para_cites.append(Citation(idx, s, e, [a], False, t[s:e]))
                i += 1
            for c in para_cites:
                if c.status == STATUS_PENDING and any(
                        fs < c.end and c.start < fe for fs, fe in in_field):
                    c.status = STATUS_AMBIGUOUS
                    c.note = ("Citation is part of a field result (e.g. an "
                              "EndNote/Zotero/Mendeley citation or a "
                              "cross-reference); left unchanged so the field "
                              "keeps working")
            citations.extend(para_cites)
            # Markers inside hyperlinks: report as AMBIGUOUS, never modify.
            for hl in p._p.iter(qn("w:hyperlink")):
                hl_text = "".join(tt.text or "" for tt in hl.iter(qn("w:t")))
                for m in self.TOKEN_RE.finditer(hl_text):
                    citations.append(Citation(
                        idx, -1, -1, [int(m.group()[1:-1])], False, m.group(),
                        status=STATUS_AMBIGUOUS,
                        note="Citation inside hyperlink; left unchanged to preserve the link"))
        return citations
