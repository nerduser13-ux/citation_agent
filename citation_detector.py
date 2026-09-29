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
never modified. Markers that span runs are resolved as AMBIGUOUS at processing
time. Reversed ranges ([3]-[1]) are reported AMBIGUOUS rather than guessed.

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
        for idx, p in enumerate(paragraphs):
            if idx >= ref_section_start:
                break
            t = direct_run_text_of(p._p)
            tokens = [
                (m.start(), m.end(), int(m.group()[1:-1]))
                for m in self.TOKEN_RE.finditer(t)
            ]
            i = 0
            while i < len(tokens):
                s, e, a = tokens[i]
                if i + 1 < len(tokens):
                    s2, e2, b = tokens[i + 1]
                    if self.RANGE_GAP_RE.match(t[e:s2]):
                        if a <= b:
                            citations.append(Citation(
                                idx, s, e2, list(range(a, b + 1)), True, t[s:e2]))
                        else:
                            citations.append(Citation(
                                idx, s, e2, [a], True, t[s:e2],
                                status=STATUS_AMBIGUOUS,
                                note="Reversed range (start > end); left unchanged"))
                        i += 2
                        continue
                citations.append(Citation(idx, s, e, [a], False, t[s:e]))
                i += 1
            # Markers inside hyperlinks: report as AMBIGUOUS, never modify.
            for hl in p._p.iter(qn("w:hyperlink")):
                hl_text = "".join(tt.text or "" for tt in hl.iter(qn("w:t")))
                for m in self.TOKEN_RE.finditer(hl_text):
                    citations.append(Citation(
                        idx, -1, -1, [int(m.group()[1:-1])], False, m.group(),
                        status=STATUS_AMBIGUOUS,
                        note="Citation inside hyperlink; left unchanged to preserve the link"))
        return citations
