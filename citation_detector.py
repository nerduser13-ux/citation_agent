"""citation_detector.py

Finds numbered citation markers in the document BODY only (the reference
section itself is excluded). Supported marker syntax:

    [1]
    [1,2]  [1, 2]  [1;2]   (several references in one bracket)
    [1-3]  [1–3]  [1, 3-5] (ranges inside one bracket)
    [1] and [2]
    [1], [2], and [3]
    [1]-[3]      (hyphen; also Unicode hyphens and the minus sign)
    [1]–[3]      (en dash)
    [1]—[3]      (em dash)
    [2] ... [2]  (repeated citations -> one Citation object per occurrence)

Ranges expand correctly: [4]-[6] and [4-6] -> 4, 5, 6 (three footnote
references), [1, 3-5] -> 1, 3, 4, 5. A citation with several numbers gets its
footnote references side by side, and citations separated only by spaces and
a comma or semicolon ("[1], [2]", "[1][2]") are placed side by side as one
group (see word_footnotes.process_paragraph and JOIN_GAP_RE).

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
ranges ([3]-[1], [3-1]) and implausibly long ones are reported AMBIGUOUS
rather than guessed.

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

# Spaces that may appear inside a marker or between two markers of one group
# (ordinary, no-break, thin, ... - but no tabs or line breaks).
_SPACES = "[ \u00a0\u2000-\u200a\u202f\u205f\u3000]*"
# Hyphen-minus, Unicode hyphen / non-breaking hyphen / figure dash, en dash,
# em dash, horizontal bar, minus sign, small and full-width hyphen-minus:
# text pasted from PDFs and web pages uses all of them.
_DASH = "[-\u2010-\u2015\u2212\ufe63\uff0d]"
_ITEM = rf"\d+(?:{_SPACES}{_DASH}{_SPACES}\d+)?"
# [1]  [1,2]  [1, 2]  [1;2]  [1-3]  [1–3]  [1, 3-5]
MARKER_RE = re.compile(rf"\[{_SPACES}{_ITEM}(?:{_SPACES}[,;]{_SPACES}{_ITEM})*{_SPACES}\]")
# What may separate two citations that are placed side by side as one group:
# spaces and at most one comma or semicolon ("[1], [2]", "[1]; [2]", "[1][2]").
# Never words ("[1] and [2]" keeps its "and").
JOIN_GAP_RE = re.compile(rf"{_SPACES}[,;]?{_SPACES}")
MAX_RANGE = 100          # longer "ranges" are surely not citations


def parse_marker(text):
    """'[1, 3-5]' -> (numbers, is_range, problem): ([1, 3, 4, 5], True, '').
    ``problem`` is a reason to leave the marker unchanged ('' if none)."""
    numbers, is_range = [], False
    for item in re.split("[,;]", text.strip()[1:-1]):
        ends = [int(n) for n in re.findall(r"\d+", item)]
        if len(ends) == 2:
            a, b = ends
            if a > b:
                return [a], True, "Reversed range (start > end); left unchanged"
            if b - a >= MAX_RANGE:
                return [a], True, "Range too long to be a citation; left unchanged"
            numbers.extend(range(a, b + 1))
            is_range = True
        else:
            numbers.extend(ends)
    return numbers, is_range, ""


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
    marker_text: str         # exact matched text, e.g. "[1]", "[1,2]" or "[4]-[6]"
    status: str = STATUS_PENDING
    note: str = ""
    fids: list = field(default_factory=list)    # (number, footnote_id) pairs
    # Set when processing: the text between the previous citation and this one
    # (only spaces and a comma/semicolon) was removed so both citations'
    # footnote references stand side by side. join_from = previous citation's end.
    joined: bool = False
    join_from: int = -1
    host = None              # run element containing the marker (set later)
    pieces = ()              # (run, lo, hi): the marker's text in earlier runs
    local_start: int = 0
    local_end: int = 0


class CitationDetector:
    TOKEN_RE = MARKER_RE
    # Separator between two range tokens: optional whitespace, one dash
    # (hyphen, en dash or em dash), optional whitespace.
    RANGE_GAP_RE = re.compile(rf"^\s*{_DASH}\s*$")

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
            tokens = [(m.start(), m.end()) + parse_marker(m.group())
                      for m in self.TOKEN_RE.finditer(t)]

            def single(tok):            # a plain [n], usable in "[a]-[b]"
                return not tok[3] and not tok[4] and len(tok[2]) == 1

            para_cites = []
            i = 0
            while i < len(tokens):
                s, e, nums, is_range, problem = tokens[i]
                if (i + 1 < len(tokens) and single(tokens[i]) and single(tokens[i + 1])
                        and self.RANGE_GAP_RE.match(t[e:tokens[i + 1][0]])):
                    e2 = tokens[i + 1][1]
                    a, b = nums[0], tokens[i + 1][2][0]
                    if a > b:
                        problem = "Reversed range (start > end); left unchanged"
                    elif b - a >= MAX_RANGE:
                        problem = "Range too long to be a citation; left unchanged"
                    para_cites.append(Citation(
                        idx, s, e2, [a] if problem else list(range(a, b + 1)), True,
                        t[s:e2], status=STATUS_AMBIGUOUS if problem else STATUS_PENDING,
                        note=problem))
                    i += 2
                    continue
                para_cites.append(Citation(
                    idx, s, e, nums, is_range, t[s:e],
                    status=STATUS_AMBIGUOUS if problem else STATUS_PENDING, note=problem))
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
                    nums, is_range, _problem = parse_marker(m.group())
                    citations.append(Citation(
                        idx, -1, -1, nums, is_range, m.group(),
                        status=STATUS_AMBIGUOUS,
                        note="Citation inside hyperlink; left unchanged to preserve the link"))
        return citations
