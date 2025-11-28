# utils/docx_handler.py
"""
Compatibility DOCX handler.

Converts a .docx to a temporary PDF (using docx2pdf or LibreOffice if available),
then delegates annotation to pdf_handler.annotate_pdf_and_build_combined (or a compatible function).

This module accepts many kwargs to remain compatible with various app.py call patterns.
"""

from pathlib import Path
import tempfile
import shutil
import subprocess
import os
import sys
from typing import Tuple, List, Optional

# Try to import pdf_handler at module import time (app.py may register it in sys.modules)
try:
    import pdf_handler  # type: ignore
except Exception:
    # If unavailable now, we'll attempt to import later inside the function into a local var.
    pdf_handler = None  # type: ignore


def _convert_with_docx2pdf(docx_path: Path, out_pdf_path: Path) -> bool:
    """
    Try conversion using docx2pdf (Python package).
    Returns True on success, False otherwise.
    """
    try:
        # Import here to avoid requiring docx2pdf unless needed
        from docx2pdf import convert  # type: ignore
        # docx2pdf.convert(src, dst) -> may overwrite, some versions accept folder mapping
        convert(str(docx_path), str(out_pdf_path))
        return out_pdf_path.exists()
    except Exception:
        return False


def _convert_with_libreoffice(docx_path: Path, out_pdf_path: Path) -> bool:
    """
    Try conversion using LibreOffice's soffice CLI.
    Example:
      soffice --headless --convert-to pdf --outdir <outdir> <docx_path>
    Returns True on success.
    """
    # Find soffice on PATH (try common names)
    soffice_cmd = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_cmd:
        return False

    outdir = out_pdf_path.parent
    try:
        cmd = [
            soffice_cmd,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(outdir),
            str(docx_path),
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        expected = outdir / (docx_path.stem + ".pdf")
        if expected.exists():
            # Move to exact destination if different
            if expected.resolve() != out_pdf_path.resolve():
                shutil.move(str(expected), str(out_pdf_path))
            return out_pdf_path.exists()
        return False
    except Exception:
        return False


def convert_docx_to_pdf(docx_path: str, output_pdf_path: str) -> None:
    """
    Convert docx -> pdf saving to output_pdf_path.
    Raises RuntimeError if conversion failed or no converter found.
    """
    src = Path(docx_path)
    dst = Path(output_pdf_path)
    dst.parent.mkdir(parents=True, exist_ok=True)

    # Try docx2pdf first
    ok = _convert_with_docx2pdf(src, dst)
    if ok:
        return

    # Next try LibreOffice / soffice
    ok = _convert_with_libreoffice(src, dst)
    if ok:
        return

    # Nothing available
    raise RuntimeError(
        "No .docx -> .pdf converter found. Install 'docx2pdf' (pip) or LibreOffice "
        "and ensure 'soffice' is on your PATH."
    )


def process_word_file(
    path: str,
    query: str,
    out_dir: str,
    resolution: int = 150,
    outline_width: int = 3,
    stroke_color: str = "red",
    force_render: bool = False,
    prefer_fallback: bool = True,
    **kwargs
) -> Tuple[List[str], Optional[str]]:
    """
    Convert DOCX → PDF then annotate via pdf_handler.

    Parameters accepted for compatibility:
      - path (str): input .docx
      - query (str): search text
      - out_dir (str): output directory
      - resolution, outline_width, stroke_color, force_render, prefer_fallback, etc.

    Returns:
      (pages, combined_pdf_path)
    """
    src = Path(path)
    if not src.exists():
        raise FileNotFoundError(f"Input DOCX not found: {path}")

    out_dir_path = Path(out_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)

    # Convert in a temporary directory
    with tempfile.TemporaryDirectory() as tmpd:
        tmpd_path = Path(tmpd)
        temp_pdf = tmpd_path / (src.stem + ".pdf")

        # Convert docx -> pdf (raises on failure)
        convert_docx_to_pdf(str(src), str(temp_pdf))

        # Ensure pdf handler is importable — use a local variable to avoid modifying module-level name
        pdf_mod = None
        try:
            # Prefer the already-imported module if available
            if "pdf_handler" in sys.modules:
                pdf_mod = sys.modules["pdf_handler"]
            elif pdf_handler is not None:  # type: ignore
                pdf_mod = pdf_handler  # type: ignore
            else:
                import importlib

                pdf_mod = importlib.import_module("pdf_handler")
        except Exception:
            raise RuntimeError(
                "pdf_handler is not importable. Ensure pdf_handler.py is available (in project root or utils/) and importable."
            )

        # Choose delegate function from pdf_mod
        delegate_fn = None
        if hasattr(pdf_mod, "annotate_pdf_and_build_combined"):
            delegate_fn = getattr(pdf_mod, "annotate_pdf_and_build_combined")
        elif hasattr(pdf_mod, "annotate_pdf"):
            # Some variants may expose annotate_pdf; try to use a compatible wrapper
            delegate_fn = getattr(pdf_mod, "annotate_pdf")
        elif hasattr(pdf_mod, "images_to_pdf") and hasattr(pdf_mod, "annotate_pdf"):
            # fallback wrapper if both exist
            def _delegate(pdf_path, q, **dkwargs):
                ann = pdf_mod.annotate_pdf(pdf_path, q, **dkwargs)
                # images_to_pdf might produce a combined pdf from images; attempt if available
                combined = None
                try:
                    combined = pdf_mod.images_to_pdf([str(pdf_path)], out_dir=dkwargs.get("out_dir", "."))
                except Exception:
                    combined = None
                return ann or [], combined

            delegate_fn = _delegate

        if delegate_fn is None:
            raise RuntimeError(
                "pdf_handler does not expose a compatible annotate function (expected annotate_pdf_and_build_combined or annotate_pdf)."
            )

        # Build kwargs forwarded to pdf handler
        pdf_kwargs = {
            "out_dir": str(out_dir_path),
            "resolution": resolution,
            "outline_width": outline_width,
            "stroke_color": stroke_color,
            "prefer_fallback": prefer_fallback,
            "force_render": force_render,
        }

        # Try to call delegate function flexibly
        try:
            # Common signature: (pdf_path, query, **kwargs) -> (pages, combined)
            result = delegate_fn(str(temp_pdf), query, **pdf_kwargs)
        except TypeError:
            # Some older variants may expect positional args: try a positional fallback
            try:
                result = delegate_fn(str(temp_pdf), query, str(out_dir_path), resolution, outline_width, stroke_color)
            except Exception as e:
                # Re-raise a helpful message
                raise RuntimeError(
                    f"Failed to call pdf handler delegate function: {e}"
                ) from e

        # Normalize result to (pages_list, combined_path)
        pages: List[str] = []
        combined: Optional[str] = None
        if isinstance(result, tuple):
            if len(result) >= 1:
                pages = result[0] or []
            if len(result) >= 2:
                combined = result[1]
        elif isinstance(result, list):
            pages = result
        elif result is None:
            pages = []
        else:
            pages = result if isinstance(result, list) else [str(result)]

        # Ensure string paths
        pages = [str(p) for p in pages] if pages else []
        combined = str(combined) if combined else None

        return pages, combined
