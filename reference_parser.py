"""reference_parser.py

Locates the document's reference / bibliography section and extracts the
number -> URL mapping. No URLs are invented: only URLs explicitly present in a
reference entry are used.
"""
import re

from document_reader import is_heading


class ReferenceParser:
    # A reference entry begins with "1." or "[1]" (or similar numbered forms).
    ENTRY_NUM_RE = re.compile(r"^\s*(?:(\d+)|\[(\d+)\])\.?\s")
    # Section-heading keywords (case-insensitive, word-boundary).
    KEYWORD_RE = re.compile(
        r"\b(references|bibliography|works\s+cited|reference\s+list|"
        r"literature\s+cited|sources)\b", re.I)

    def find_reference_section(self, paragraphs, config):
        """Return (start_index, method, confidence, note) or None.

        start_index is the first paragraph *after* the heading that belongs to
        the reference list. Returns None when the section cannot be confidently
        identified -> the caller must stop safely (no guessing)."""
        texts = [p.text for p in paragraphs]
        n = len(texts)

        # 1) Explicit heading (preferred, high confidence).
        for i, (p, t) in enumerate(zip(paragraphs, texts)):
            if self.KEYWORD_RE.search(t) and (is_heading(p) or len(t.strip()) < 60):
                return i + 1, "heading", "high", (
                    f'Reference heading detected: "{t.strip()[:60]}"')

        # 2) Fallback: largest contiguous block of numbered entries near the end
        #    that also contains URLs. Medium confidence; reported.
        best = None
        i = 0
        while i < n:
            if self.ENTRY_NUM_RE.match(texts[i]):
                j = i
                while j < n and (
                        self.ENTRY_NUM_RE.match(texts[j])
                        or (texts[j].strip() and not is_heading(paragraphs[j]))):
                    j += 1
                block = texts[i:j]
                entries = sum(1 for x in block if self.ENTRY_NUM_RE.match(x))
                urls = sum(1 for x in block if "http" in x.lower())
                if (entries >= config.min_reference_entries
                        and urls >= config.min_reference_urls):
                    if best is None or entries > best[2]:
                        best = (i, j, entries, urls)
                i = j
            else:
                i += 1
        if best:
            return best[0], "block-detection", "medium", (
                f"Inferred reference block ({best[2]} entries) without explicit "
                f"heading")
        return None

    def parse_references(self, paragraphs, start_index):
        """Extract number -> {url, text} for every reference entry beginning at
        start_index. Stops at the next clear heading."""
        entries = []
        current = None
        for p in paragraphs[start_index:]:
            t = p.text
            if is_heading(p) and not self.ENTRY_NUM_RE.match(t):
                break
            m = self.ENTRY_NUM_RE.match(t)
            if m:
                if current:
                    entries.append(current)
                num = int(m.group(1) or m.group(2))
                current = {"num": num, "text": t}
            else:
                if current is not None and t.strip():
                    current["text"] += " " + t.strip()
        if current:
            entries.append(current)

        ref_map = {}
        for e in entries:
            urls = re.findall(r"https?://[^\s\)\]]+", e["text"])
            ref_map[e["num"]] = {
                "url": self._pick_url(urls),
                "text": e["text"],
            }
        return ref_map

    @staticmethod
    def _pick_url(urls):
        """Prefer a direct source URL over a bare DOI resolver when both exist."""
        if not urls:
            return None
        non_doi = [u for u in urls if "doi.org" not in u.lower()]
        if non_doi:
            return non_doi[-1]
        return urls[-1]
