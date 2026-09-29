"""reference_parser.py

Locates the document's reference / bibliography section and extracts the
number -> URL mapping. No URLs are invented: only URLs explicitly present in a
reference entry are used (bare "DOI: 10.1000/abc" text is NOT turned into a URL).

Detection strategy (most conservative guess wins; ambiguity aborts):

1. Heading with a *styled* heading (Heading 1..9 / Title / outline level) whose
   text, after stripping leading numbering ("5.", "1.2", "A.") and trailing
   punctuation, is exactly a keyword ("references", "bibliography", ...).
   -> confidence "high", method "heading".
2. Same, but heading identified only by the bold+short heuristic.
   -> confidence "high", method "heading-bold".
3. A heading that merely *contains* a keyword ("Data Sources"), accepted only if
   it is directly followed by a qualifying numbered block.
   -> confidence "medium", method "heading+block".
4. No heading at all: a numbered block near the end of the document
   (starts at/after ``config.block_start_fraction`` of the paragraph list, at
   least ``config.min_reference_entries`` numbered entries, at least
   ``config.min_reference_urls`` URLs in entry lines).
   -> confidence "medium", method "block-detection".
5. Nothing else: None -> the caller must abort safely (no partial output).

When several candidates exist, the LAST one in document order wins (reference
sections sit at the end of the document by convention). Paragraphs inside
table cells are never heading candidates (a bold "Reference" column header in
an appendix table must not take over).

The reference list ends at the next styled heading or bold-only heading line
(see ``parse_references``), so a following section's URLs are never
attributed to the last reference.

URL selection rule when an entry contains several URLs (deterministic,
documented): the LAST URL that is not a DOI resolver (no "doi.org" in the URL);
if every URL is a DOI resolver, the LAST URL. URLs are used exactly as written,
except a single run of trailing sentence punctuation (.,;:!? ) is stripped.
"""
import re

from document_reader import in_table, is_heading, is_style_heading, raw_text_of
from docx.oxml.ns import qn

# A reference entry line begins with "1.", "1)", "1 ", or "[1]" (then a space).
ENTRY_NUM_RE = re.compile(r"^\s*(?:\[(\d+)\]|(\d+)[\.\)])\s+")
# URL: http(s) scheme, no whitespace/brackets/quotes.
URL_RE = re.compile(r"https?://[^\s\(\)\[\]<>\u201c\u201d\"']+")
# Leading numbering such as "5.", "1.2.", "5)", "(3)", "A.".
_NUM_PREFIX_RE = re.compile(r"^\s*(?:(?:\d+[\.\)])|[\(（]\d+[\)）])+[\.\)]?\s*")
_LETTER_PREFIX_RE = re.compile(r"^\s*[A-Z][\.\)]\s+")
YEAR_RANGE = range(1900, 2100)
_YEAR_START_RE = re.compile(r"^\s*(19\d{2}|20\d{2})\.\s")


def _normalize_heading(text: str) -> str:
    s = _NUM_PREFIX_RE.sub("", text.strip())
    s = _LETTER_PREFIX_RE.sub("", s)
    return s.rstrip(":\u2013\u2014- \t").strip().lower()


def _is_entry_line(text: str):
    """Return the entry number if the line starts a numbered reference entry,
    else None. A leading year ("2020. Author ...") is NOT an entry number."""
    m = ENTRY_NUM_RE.match(text)
    if not m:
        return None
    num = int(m.group(1) or m.group(2))
    # "2020." style year-start -> not a reference number.
    stripped = text.lstrip()
    if (m.group(2) is not None
            and num in YEAR_RANGE
            and stripped.startswith(f"{num}.")):
        return None
    return num


class ReferenceParser:
    def __init__(self, config=None):
        self.config = config
        self.warnings = []   # human-readable notes accumulated while parsing

    # ------------------------------------------------------------------
    # Section location
    # ------------------------------------------------------------------
    def find_reference_section(self, paragraphs, config):
        """Return (start_index, method, confidence, note) or None.

        start_index is the first paragraph of the reference list (the first
        entry line, or the paragraph after the heading). None means the
        section cannot be confidently identified -> abort, do not guess."""
        cfg = config or self.config
        n = len(paragraphs)
        if n == 0:
            return None

        keywords = tuple(k.lower() for k in cfg.reference_keywords)

        # --- 1+2: exact keyword headings (styled preferred) -------------
        # Table cells are never section headings: a bold "Reference" /
        # "Sources" column header (e.g. a literature matrix in an appendix)
        # must not be mistaken for the reference-section heading.
        candidates = []          # (index, method)
        for i, p in enumerate(paragraphs):
            text = (p.text or "").strip()
            if not text or in_table(p):
                continue
            if not (is_style_heading(p) or is_heading(p)):
                continue
            norm = _normalize_heading(text)
            if norm in keywords:
                method = "heading" if is_style_heading(p) else "heading-bold"
                candidates.append((i, method))
        if candidates:
            idx, method = candidates[-1]
            return idx + 1, method, "high", (
                f'Reference heading detected: "{(paragraphs[idx].text or "").strip()[:60]}"')

        # --- 3: keyword-containing heading + qualifying block ------------
        for i, p in enumerate(paragraphs):
            text = (p.text or "").strip()
            if not text or not is_style_heading(p) or in_table(p):
                continue
            norm = _normalize_heading(text)
            if not any(k in norm for k in keywords):
                continue
            # First non-blank paragraph after the heading must start the block
            # (a styled heading there means the "heading" is a section title
            # and its content is unrelated).
            j = i + 1
            while j < len(paragraphs) and not (paragraphs[j].text or "").strip():
                j += 1
            if j < len(paragraphs) and is_style_heading(paragraphs[j]):
                continue
            block = self._scan_block(paragraphs, j)
            if block and block.entries >= cfg.min_reference_entries \
                    and block.urls >= cfg.min_reference_urls:
                return block.start, "heading+block", "medium", (
                    f'Heading "{text[:60]}" followed by a numbered reference block '
                    f'({block.entries} entries)')

        # --- 4: numbered block near the end of the document --------------
        min_start = int(n * cfg.block_start_fraction)
        best = None
        i = 0
        while i < n:
            if i >= min_start and _is_entry_line((paragraphs[i].text or "")):
                block = self._scan_block(paragraphs, i)
                if (block and block.entries >= cfg.min_reference_entries
                        and block.urls >= cfg.min_reference_urls):
                    key = (block.entries, block.urls, block.start)
                    if best is None or key > best[0]:
                        best = (key, block)
                i = block.end if block else i + 1
            else:
                i += 1
        if best:
            block = best[1]
            return block.start, "block-detection", "medium", (
                f"Inferred reference block ({block.entries} entries, "
                f"{block.urls} URLs) near end of document, no explicit heading")
        return None

    # ------------------------------------------------------------------
    # Block scanning
    # ------------------------------------------------------------------
    def _scan_block(self, paragraphs, start):
        """Maximal run of paragraphs starting at ``start`` that forms a
        candidate reference list. Returns a small object or None."""
        if start >= len(paragraphs):
            return None
        if not _is_entry_line((paragraphs[start].text or "")):
            return None
        entries = 0
        urls = 0
        j = start
        while j < len(paragraphs):
            p = paragraphs[j]
            if is_style_heading(p):
                break
            t = (p.text or "").strip()
            if _is_entry_line(t):
                entries += 1
                urls += len(_extract_urls(t))
            elif is_heading(p):
                break   # bold-only heading: the numbered block is over
            j += 1
        return _Block(start, j, entries, urls)

    # ------------------------------------------------------------------
    # Entry parsing
    # ------------------------------------------------------------------
    def parse_references(self, paragraphs, start_index):
        """Extract number -> {url, text} for the reference entries from
        ``start_index`` to the next heading (or end of document).

        The list ends at the next styled heading, or at a bold-only heading
        line (short, entirely bold, not itself a numbered entry) - documents
        formatted with direct bold instead of heading styles would otherwise
        glue a following section (e.g. "Appendix A") onto the last entry and
        attribute that section's URLs to the last reference.

        Rules:
          * continuation (indented/wrapped) lines are appended to the current
            entry;
          * a line that starts with a year ("2020. ...") is NOT a numbered
            entry and NOT a continuation: it ends the current entry (its text
            and URLs must not be attributed to a reference number);
          * the FIRST occurrence of a duplicate number wins (duplicates are
            recorded in self.warnings);
          * only explicit http(s) URLs are used - nothing is fabricated.

        Returns (ref_map, warnings)."""
        self.warnings = []
        entries = {}      # num -> {"text": str}
        current = None
        list_counters = {}
        for p in paragraphs[start_index:]:
            if is_style_heading(p):
                break
            t = raw_text_of(p._p)
            num = _is_entry_line(t)
            if num is None:
                label = self._automatic_number(p, list_counters)
                if label is not None:
                    num, t = label, f"{label}. {t}"
            if num is None and is_heading(p):
                break   # bold-only heading: the reference list is over
            if num is not None:
                if num in entries:
                    self.warnings.append(
                        f"Duplicate reference number {num}; first occurrence kept")
                    current = None
                    continue
                current = num
                entries[num] = t.strip()
            elif _YEAR_START_RE.match(t) and t.strip():
                if current is not None:
                    self.warnings.append(
                        f"Line starting with a year is not a numbered entry; "
                        f"it ends reference {current}: {t.strip()[:40]}...")
                current = None
            elif t.strip():
                if current is not None:
                    entries[current] += " " + t.strip()
        ref_map = {}
        for num in sorted(entries):
            urls = _extract_urls(entries[num])
            ref_map[num] = {
                "url": _pick_url(urls),
                "text": entries[num],
            }
        return ref_map, self.warnings

    @staticmethod
    def _automatic_number(paragraph, counters):
        """Read decimal Word list labels that are absent from paragraph text."""
        ppr = paragraph._p.pPr
        num_pr = ppr.find(qn("w:numPr")) if ppr is not None else None
        if num_pr is None:
            return None
        num_id_el = num_pr.find(qn("w:numId"))
        if num_id_el is None:
            return None
        num_id = num_id_el.get(qn("w:val"))
        level_el = num_pr.find(qn("w:ilvl"))
        ilvl = int(level_el.get(qn("w:val"), "0")) if level_el is not None else 0
        numbering = paragraph.part.numbering_part.element
        num = next((x for x in numbering.findall(qn("w:num"))
                    if x.get(qn("w:numId")) == num_id), None)
        if num is None:
            return None
        abstract_id = num.find(qn("w:abstractNumId"))
        if abstract_id is None:
            return None
        abstract = next((x for x in numbering.findall(qn("w:abstractNum"))
                         if x.get(qn("w:abstractNumId")) == abstract_id.get(qn("w:val"))), None)
        level = None
        if abstract is not None:
            level = next((x for x in abstract.findall(qn("w:lvl"))
                          if int(x.get(qn("w:ilvl"), "0")) == ilvl), None)
        if level is None:
            return None
        fmt = level.find(qn("w:numFmt"))
        if fmt is None or fmt.get(qn("w:val")) != "decimal":
            return None
        key = (num_id, ilvl)
        start = level.find(qn("w:start"))
        counters[key] = counters.get(key, int(start.get(qn("w:val"), "1")) - 1 if start is not None else 0) + 1
        return counters[key]


class _Block:
    __slots__ = ("start", "end", "entries", "urls")

    def __init__(self, start, end, entries, urls):
        self.start = start
        self.end = end
        self.entries = entries
        self.urls = urls


def _extract_urls(text: str):
    """Explicit URLs in ``text``; trailing sentence punctuation stripped."""
    out = []
    for u in URL_RE.findall(text):
        u = u.rstrip(".,;:!?")
        if u:
            out.append(u)
    return out


def _pick_url(urls):
    """Documented selection rule: last non-DOI-resolver URL, else last URL."""
    if not urls:
        return None
    non_doi = [u for u in urls if "doi.org" not in u.lower()]
    return non_doi[-1] if non_doi else urls[-1]
