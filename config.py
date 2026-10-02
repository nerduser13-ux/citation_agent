"""Configuration for the Word Citation Footnote Agent.

All paths, behaviour switches and detection thresholds live here so the tool can be
driven from the CLI, an AI agent, or a future GUI without touching the logic.
"""
from pathlib import Path


class Config:
    # --- I/O -----------------------------------------------------------------
    input_path: str = None          # explicit input .docx (CLI --input)
    output_path: str = None         # explicit output .docx (CLI --output)
    input_dir: Path = Path("input")
    output_dir: Path = Path("output")

    # --- Footnote backend -----------------------------------------------------
    # "xml"  -> cross-platform OOXML manipulation (default, validated here)
    # "com"  -> Windows Microsoft Word COM automation (experimental)
    backend: str = "xml"

    # --- Marker handling ------------------------------------------------------
    # True  -> the [n] marker is REPLACED by the Word footnote reference marker.
    # False -> the [n] marker is KEPT and the footnote reference is inserted
    #          immediately AFTER it (the [n] text survives).
    replace_marker: bool = True

    # Text between footnote references that stand side by side - for [1,2],
    # [1-3] or "[1], [2]" - formatted like the references (superscript):
    # " " -> 1 2,  "," -> 1,2,  "" -> 12 (adjacent; reads like twelve).
    footnote_separator: str = " "

    # If the resolved output file already exists, create a safe alternative
    # (e.g. document_with_footnotes_1.docx) unless this is True.
    overwrite_output: bool = False

    # --- Set by main.run() ----------------------------------------------------
    resolved_output_path: Path = None   # output path actually used
    report_path: Path = None            # CSV review report actually written

    # --- Reference-section detection ------------------------------------------
    # Keyword phrases a reference-section heading may consist of (case-insensitive).
    reference_keywords: tuple = (
        "references", "reference", "bibliography", "works cited",
        "reference list", "literature cited", "sources",
    )
    # Fallback "numbered block near the end of the document" thresholds.
    min_reference_entries: int = 3   # minimum numbered entries in a candidate block
    min_reference_urls: int = 1      # minimum URLs found in entry lines of a block
    block_start_fraction: float = 0.40
    # A candidate block must start no earlier than this fraction of the way through
    # the document's paragraph list (references are near the end, by convention).

    def resolve_output_path(self):
        """Return a safe output path, never overwriting the input or an existing
        output unless explicitly allowed."""
        src = Path(self.input_path)
        if self.output_path:
            out = Path(self.output_path)
        else:
            out = Path(self.output_dir) / (src.stem + "_with_footnotes" + src.suffix)
        if out.resolve() == src.resolve():
            raise ValueError(
                "Output path is identical to the input path; refusing to touch the "
                "original file. Choose a different --output (or use the default).")
        if out.exists() and not self.overwrite_output:
            i = 1
            while True:
                cand = out.with_name(f"{out.stem}_{i}{out.suffix}")
                if not cand.exists():
                    out = cand
                    break
                i += 1
        out.parent.mkdir(parents=True, exist_ok=True)
        return out
