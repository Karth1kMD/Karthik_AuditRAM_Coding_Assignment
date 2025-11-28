#!/usr/bin/env python3
"""
app.py

Unified runner for the annotation pipeline with robust dynamic imports.

This version also attempts to load pdf_handler.py from either:
 - project root (ROOT/pdf_handler.py)
 - utils folder (ROOT/utils/pdf_handler.py)

so that `import pdf_handler` inside utils modules works without editing handler files.
"""

import sys
import os
import argparse
from pathlib import Path
import subprocess
import platform
import traceback
import importlib
import importlib.util

ROOT = Path(__file__).resolve().parent
UTILS_DIR = ROOT / "utils"

# -------------------------
# Robust import helpers
# -------------------------
def import_module_by_path(module_name: str, filepath: Path, verbose: bool = True):
    """Load a module from a file path, register it in sys.modules and return the module object or None."""
    try:
        spec = importlib.util.spec_from_file_location(module_name, str(filepath))
        module = importlib.util.module_from_spec(spec)
        # Insert into sys.modules before executing to allow intra-module imports (circular safe)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        if verbose:
            print(f"[app.py] Loaded module {module_name} from {filepath}")
        return module
    except Exception:
        # Print debug so user sees why the import failed
        print(f"[app.py] DEBUG: failed loading {module_name} from {filepath}")
        traceback.print_exc()
        if module_name in sys.modules:
            del sys.modules[module_name]
        return None

def try_import(name: str):
    """Try import by name, return module or None (silently)."""
    try:
        return importlib.import_module(name)
    except Exception:
        return None

# -------------------------
# Load handlers robustly
# -------------------------
# 1) Try normal imports first
pdf_handler = try_import("pdf_handler")
docx_handler = try_import("utils.docx_handler")
excel_handler = try_import("utils.excel_handler")
image_handler = try_import("utils.image_handler")

# 2) If any missing, ensure project root is on sys.path
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# 3) If pdf_handler missing, try to load from project root OR utils folder (register as 'pdf_handler')
if pdf_handler is None:
    candidates = [
        ROOT / "pdf_handler.py",
        UTILS_DIR / "pdf_handler.py"
    ]
    for cand in candidates:
        if cand.exists():
            pdf_mod = import_module_by_path("pdf_handler", cand)
            if pdf_mod:
                pdf_handler = pdf_mod
                break

# 4) Load utils modules (docx/excel/image). Prefer package import, otherwise load by path.
mapping = {
    "utils.docx_handler": UTILS_DIR / "docx_handler.py",
    "utils.excel_handler": UTILS_DIR / "excel_handler.py",
    "utils.image_handler": UTILS_DIR / "image_handler.py",
}
for mod_name, path in mapping.items():
    # skip if already importable
    if try_import(mod_name) is not None:
        # attach to locals by name
        mod = importlib.import_module(mod_name)
        if mod_name.endswith("docx_handler"):
            docx_handler = mod
        elif mod_name.endswith("excel_handler"):
            excel_handler = mod
        elif mod_name.endswith("image_handler"):
            image_handler = mod
        continue

    # otherwise try load by path
    if path.exists():
        loaded = import_module_by_path(mod_name, path)
        if loaded:
            if mod_name.endswith("docx_handler"):
                docx_handler = loaded
            elif mod_name.endswith("excel_handler"):
                excel_handler = loaded
            elif mod_name.endswith("image_handler"):
                image_handler = loaded

# 5) Friendly warning showing what's missing
_missing = []
if pdf_handler is None:
    _missing.append("pdf_handler (pdf_handler.py in root or utils/)")
if docx_handler is None:
    _missing.append("utils.docx_handler (utils/docx_handler.py)")
if excel_handler is None:
    _missing.append("utils.excel_handler (utils/excel_handler.py)")
if image_handler is None:
    _missing.append("utils.image_handler (utils/image_handler.py)")

if _missing:
    print("[app.py] WARNING: Some handlers failed to import:", ", ".join(_missing))
    print("[app.py] If you see errors later, ensure these files exist and are valid Python modules.")
    print("[app.py] Current working dir:", Path.cwd())
    print("[app.py] sys.path[0]:", sys.path[0])
    print()

# -------------------------
# Dispatcher wrappers
# -------------------------
def detect_type(path: Path):
    SUPPORTED_EXTENSIONS = {
        ".pdf": "pdf",
        ".docx": "docx",
        ".xlsx": "excel",
        ".xls": "excel",
        ".png": "image",
        ".jpg": "image",
        ".jpeg": "image",
        ".tif": "image",
        ".tiff": "image",
    }
    return SUPPORTED_EXTENSIONS.get(path.suffix.lower())

def open_folder(path: Path):
    try:
        if platform.system() == "Windows":
            os.startfile(str(path))
        elif platform.system() == "Darwin":
            subprocess.run(["open", str(path)])
        else:
            subprocess.run(["xdg-open", str(path)])
    except Exception:
        pass

# Each runner will raise a helpful error if its handler is missing
def run_pdf(path: Path, query: str, outdir: Path, resolution: int, outline: int, color: str, force_render: bool):
    if pdf_handler is None:
        raise RuntimeError("pdf_handler not importable.")
    if hasattr(pdf_handler, "annotate_pdf_and_build_combined"):
        fn = pdf_handler.annotate_pdf_and_build_combined
        pages, combined = fn(str(path), query, out_dir=str(outdir),
                             resolution=resolution, outline_width=outline,
                             stroke_color=color, prefer_fallback=True, force_render=force_render)
        return {"pages": pages, "combined": str(combined) if combined else None}
    else:
        raise RuntimeError("pdf_handler imported but expected function not found: annotate_pdf_and_build_combined")

def run_docx(path: Path, query: str, outdir: Path, resolution: int, outline: int, color: str, force_render: bool):
    if docx_handler is None:
        raise RuntimeError("utils.docx_handler not importable.")
    if hasattr(docx_handler, "process_word_file"):
        pages, combined = docx_handler.process_word_file(str(path), query, out_dir=str(outdir),
                                                         resolution=resolution, outline_width=outline,
                                                         stroke_color=color, force_render=force_render,
                                                         prefer_fallback=True)
        return {"pages": pages, "combined": str(combined) if combined else None}
    else:
        raise RuntimeError("docx_handler imported but expected function not found: process_word_file")

def run_excel(path: Path, query: str, outdir: Path, outline: int, color: str, font_size: int):
    if excel_handler is None:
        raise RuntimeError("utils.excel_handler not importable.")
    if hasattr(excel_handler, "process_excel_file"):
        images, combined, annotated_xlsx = excel_handler.process_excel_file(
            str(path),
            query,
            out_dir=str(outdir),
            stroke_color=color,
            outline_width=outline,
            font_size=font_size
        )
        return {"pages": images, "combined": combined, "annotated_xlsx": annotated_xlsx}
    else:
        raise RuntimeError("excel_handler imported but expected function not found: process_excel_file")

def run_image(path: Path, query: str, outdir: Path, outline: int, color: str):
    if image_handler is None:
        raise RuntimeError("utils.image_handler not importable.")
    if hasattr(image_handler, "process_image"):
        img, combined = image_handler.process_image(
            str(path),
            query,
            out_dir=str(outdir),
            stroke_color=color,
            outline_width=outline
        )
        return {"pages": [img], "combined": combined}
    else:
        raise RuntimeError("image_handler imported but expected function not found: process_image")

DISPATCH = {
    "pdf": run_pdf,
    "docx": run_docx,
    "excel": run_excel,
    "image": run_image
}

# -------------------------
# Utility functions
# -------------------------
def process_single_file(file_path: Path, query: str, outdir: Path, opts: dict):
    ftype = detect_type(file_path)
    if not ftype:
        raise RuntimeError(f"Unsupported extension: {file_path.suffix}")
    out_subdir = outdir / file_path.stem
    out_subdir.mkdir(parents=True, exist_ok=True)

    runner = DISPATCH.get(ftype)
    if runner is None:
        raise RuntimeError(f"No runner available for type: {ftype}")

    result = runner(file_path, query, out_subdir, **opts)
    return result, out_subdir

def gather_files(path: Path, recursive: bool):
    if path.is_file():
        return [path]
    files = []
    for p in path.iterdir():
        if p.is_file():
            if detect_type(p) is not None:
                files.append(p)
        elif recursive and p.is_dir():
            files.extend(gather_files(p, recursive))
    return files

def interactive_prompt():
    print("\n=== AuditRAM Annotation Runner (interactive mode) ===\n")
    while True:
        inp = input("Enter path to file or folder (press Enter to cancel): ").strip()
        if inp == "":
            print("Cancelled by user.")
            sys.exit(0)
        p = Path(inp).expanduser().resolve()
        if not p.exists():
            print("Path does not exist. Try again.")
            continue
        break

    while True:
        q = input("Enter search keyword/text (case-insensitive): ").strip()
        if q == "":
            print("Please enter a non-empty keyword.")
            continue
        break

    return p, q

# -------------------------
# Main
# -------------------------
def main():
    parser = argparse.ArgumentParser(description="Unified annotation runner (interactive if no args provided)")
    parser.add_argument("path", nargs="?", help="File or directory to process (optional). If omitted the script will prompt.")
    parser.add_argument("query", nargs="?", help="Text to search for (optional). If omitted the script will prompt.")
    parser.add_argument("--outdir", default="output", help="Base output directory (default: output/)")
    parser.add_argument("--recursive", action="store_true", help="If path is a directory, scan recursively")
    parser.add_argument("--resolution", type=int, default=150, help="Raster DPI for PDF/DOCX rendering")
    parser.add_argument("--outline", type=int, default=3, help="Outline width in pixels")
    parser.add_argument("--color", default="red", help="Outline color")
    parser.add_argument("--font-size", type=int, default=14, help="Font size for Excel rendering")
    parser.add_argument("--force-render", action="store_true", help="Force rasterization even when no matches are found")
    parser.add_argument("--open", action="store_true", help="Open the output folder when finished (if supported)")

    args = parser.parse_args()

    # Interactive fallback if path or query not provided
    if not args.path or not args.query:
        path_obj, query = interactive_prompt()
        recursive = args.recursive
        outdir = Path(args.outdir).resolve()
        resolution = args.resolution
        outline = args.outline
        color = args.color
        force_render = args.force_render
        font_size = args.font_size
        open_after = args.open
    else:
        path_obj = Path(args.path).expanduser().resolve()
        query = args.query
        recursive = args.recursive
        outdir = Path(args.outdir).resolve()
        resolution = args.resolution
        outline = args.outline
        color = args.color
        force_render = args.force_render
        font_size = args.font_size
        open_after = args.open

    outdir.mkdir(parents=True, exist_ok=True)

    files = gather_files(path_obj, recursive)
    if not files:
        print("[runner] No supported files found at:", path_obj)
        sys.exit(2)

    overall = []
    for f in files:
        print("\n" + "=" * 60)
        print(f"[runner] Processing: {f.name}  ({f})")
        try:
            ftype = detect_type(f)
            if ftype == "pdf" or ftype == "docx":
                runner_opts = {
                    "resolution": resolution,
                    "outline": outline,
                    "color": color,
                    "force_render": force_render
                }
            elif ftype == "excel":
                runner_opts = {"outline": outline, "color": color, "font_size": font_size}
            elif ftype == "image":
                runner_opts = {"outline": outline, "color": color}
            else:
                runner_opts = {}

            result, subdir = process_single_file(f, query, outdir, runner_opts)
            print(f"[runner] DONE: outputs saved to {subdir}")
            if result.get("pages"):
                print("  Annotated images:")
                for p in result["pages"]:
                    print("   -", p)
            if result.get("combined"):
                print("  Combined PDF:", result["combined"])
            if result.get("annotated_xlsx"):
                print("  Annotated Excel:", result["annotated_xlsx"])

            overall.append({"file": str(f), "outdir": str(subdir), "result": result})

        except Exception as e:
            print(f"[runner] ERROR processing {f.name}: {e}")
            traceback.print_exc()
            overall.append({"file": str(f), "error": str(e)})

    print("\n" + "=" * 60)
    print("[runner] Summary:")
    for item in overall:
        if item.get("error"):
            print(" -", item["file"], "→ ERROR:", item["error"])
        else:
            r = item["result"]
            print(" -", item["file"], "→ outdir:", item["outdir"])
            if r.get("combined"):
                print("    combined:", r["combined"])
            if r.get("pages"):
                print("    pages:", len(r["pages"]))

    if open_after:
        open_folder(outdir)

    print("\n[runner] All done.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
