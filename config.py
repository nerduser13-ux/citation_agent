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
    # "com"  -> Windows Microsoft Word COM automation (experimental, untested on Linux)
    backend: str = "xml"

    # --- Marker handling -----------------------------------------------------
    # True  -> the [n] marker is REPLACED by the Word footnote reference marker
    #          (the [n] disappears, as required by the spec).
    # False -> the [n] marker is KEPT and a footnote reference is inserted after it.
    replace_marker: bool = True

    # If the resolved output file already exists, create a safe alternative
    # (e.g. document_with_footnotes_1.docx) unless this is True.
    overwrite_output: bool = False

    # --- Reference-section detection -----------------------------------------
    reference_keywords: tuple = (
        "references", "bibliography", "works cited",
        "reference list", "literature cited", "sources",
    )
    min_reference_entries: int = 3   # fallback block-detection thresholds
    min_reference_urls: int = 1

    def resolve_output_path(self):
        """Return a safe output path, never overwriting the input or an existing
        output unless explicitly allowed."""
        src = Path(self.input_path)
        if self.output_path:
            out = Path(self.output_path)
        else:
            out = self.output_dir / (src.stem + "_with_footnotes" + src.suffix)
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
