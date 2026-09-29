"""citation_detector.py

Finds numbered citation markers ([1], [1], [2], [1-3], [1] and [2], ...) in the
document BODY ONLY (the reference section is excluded). Produces lightweight
Citation objects; the actual placement/safety checks happen at insertion time.

IMPORTANT: marker offsets are measured over the paragraph's *direct* <w:r> run
text (hyperlinks excluded). This matches the coordinate system used by the
surgical inserter in word_footnotes.py, so a marker that sits after a hyperlink
is not mis-attributed to it.
"""
import re
from dataclasses import dataclass, field

from docx.oxml.ns import qn


@dataclass
class Citation:
    para_index: int          # index into the document's paragraph list
    start: int               # global character offset of the marker
    end: int                 # global character offset just past the marker
    numbers: list            # reference numbers this marker maps to
    is_range: bool           # True for [1-3] style ranges
    marker_text: str         # exact matched text, e.g. "[1]" or "[1-3]"
    status: str = "pending"  # pending | SUCCESS | ERROR | SKIPPED
    note: str = ""
    fids: list = field(default_factory=list)   # (number, footnote_id) pairs
    host = None              # run element that contains the marker (set later)
    local_start: int = 0
    local_end: int = 0


class CitationDetector:
    TOKEN_RE = re.compile(r"\[\d+\]")
    RANGE_GAP_RE = re.compile(r"^\s*[\u2013\-]\s*$")   # en/em dash or hyphen

    @staticmethod
    def _direct_text(p):
        return "".join(
            t.text or "" for r in p._p.findall(qn("w:r"))
            for t in r.findall(qn("w:t")))

    def detect(self, paragraphs, ref_section_start):
        """Detect citation markers in body paragraphs only (index < ref_section_start)."""
        citations = []
        for idx, p in enumerate(paragraphs):
            if idx >= ref_section_start:
                break
            # Coordinates must match the surgical inserter: direct runs only.
            t = self._direct_text(p)
            tokens = [
                (m.start(), m.end(), int(m.group()[1:-1]))
                for m in self.TOKEN_RE.finditer(t)
            ]
            i = 0
            while i < len(tokens):
                s, e, a = tokens[i]
                if i + 1 < len(tokens):
                    s2, e2, a2 = tokens[i + 1]
                    if self.RANGE_GAP_RE.match(t[e:s2]):
                        nums = list(range(a, a2 + 1)) if a <= a2 else [a]
                        citations.append(
                            Citation(idx, s, e2, nums, True, t[s:e2]))
                        i += 2
                        continue
                citations.append(Citation(idx, s, e, [a], False, t[s:e]))
                i += 1
            # Markers that live inside a hyperlink cannot be safely modified;
            # report them (and leave them untouched) rather than guessing.
            for hl in p._p.findall(".//" + qn("w:hyperlink")):
                hl_text = "".join(
                    tt.text or "" for tt in hl.iter(qn("w:t")))
                for m in self.TOKEN_RE.finditer(hl_text):
                    citations.append(Citation(
                        idx, -1, -1, [int(m.group()[1:-1])], False, m.group(),
                        status="ERROR",
                        note="Citation inside hyperlink; skipped to preserve link"))
        return citations
