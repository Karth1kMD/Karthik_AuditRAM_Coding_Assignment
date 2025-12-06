<!-- .github/copilot-instructions.md - Guidance for AI coding agents -->
# Repository-specific Copilot instructions

This project is a small annotation runner that finds occurrences of text inside documents (PDF/DOCX/Excel/Image)
and produces annotated images and a combined PDF. Below are the concrete, actionable conventions and patterns
an AI agent should follow when making changes.

- **Big picture**: `app.py` is the unified CLI/runner and the integration point. It dynamically imports handler
  modules (PDF/DOCX/Excel/Image) and dispatches to handler-specific functions. Handlers live under `utils/` and
  must expose well-known functions so `app.py` can call them.

- **Key files to read before changing behavior**:
  - `app.py` — orchestrator, dynamic import helpers, `DISPATCH` mapping, CLI flags.
  - `utils/pdf_handler.py` — provides `annotate_pdf_and_build_combined(pdf_path, query, ...)` and `images_to_pdf(...)`.
  - `utils/docx_handler.py` — converts DOCX→PDF then delegates to pdf handler via `process_word_file(...)`.
  - `utils/excel_handler.py` — exposes `process_excel_file(...)` and relies on `pdf_handler.images_to_pdf`.
  - `utils/image_handler.py` — exposes `process_image(...)` and uses pytesseract for OCR.
  - `tests/test_detect_type.py` — shows expected `detect_type(Path)` behavior (case-insensitive suffix mapping).

- **Handler contract / expected symbols** (do not rename without updating calls):
  - `pdf_handler.annotate_pdf_and_build_combined(pdf_path, query, out_dir=..., resolution=..., outline_width=..., stroke_color=..., prefer_fallback=True, force_render=False)` → returns `(pages_list, combined_pdf_path)`.
  - `pdf_handler.images_to_pdf(image_list, out_pdf)` → write combined PDF.
  - `docx_handler.process_word_file(path, query, out_dir, resolution=..., outline_width=..., stroke_color=..., force_render=..., prefer_fallback=...)` → returns `(pages, combined)`.
  - `excel_handler.process_excel_file(xlsx_path, query, out_dir=..., stroke_color=..., outline_width=..., font_size=...)` → returns `(generated_images, combined_pdf_path, annotated_xlsx_path)`.
  - `image_handler.process_image(image_path, query, out_dir=..., stroke_color=..., outline_width=...)` → returns `(annotated_image_path, combined_pdf_path)`.

- **Dynamic import pattern**:
  - `app.py` attempts normal imports like `utils.docx_handler` first; if missing it will load by file path and register the module name in `sys.modules`.
  - `pdf_handler.py` may exist in project root or in `utils/`. `app.py` will try both locations and register the loaded module under the name `pdf_handler`.
  - When editing handlers, preserve compatible function names/signatures or add a compatibility wrapper in `app.py`/handler to avoid breaking runtime imports.

- **Common workflows & commands**:
  - Run the annotation runner: `python app.py <path> "query text" --outdir output --recursive --resolution 150`
  - Run a handler directly (for local debugging), e.g.: `python utils/excel_handler.py sample.xlsx "query" --outdir output`
  - Run tests: `pytest -q` from the repository root. The project uses `tests/` (see `tests/test_detect_type.py`).
  - Install runtime deps (see `Requirements.txt`): `pip install -r Requirements.txt`.

- **Platform & dependency notes (Windows specifics)**:
  - `image_handler` depends on Tesseract OCR binary. On Windows, ensure Tesseract is installed and on PATH.
  - `docx_handler` needs either the `docx2pdf` Python package or LibreOffice (`soffice`) on PATH to convert .docx → .pdf.
  - `pdf_handler` uses `pdfplumber` and `Pillow` for rendering; `excel_handler` uses `openpyxl`.

- **Project-specific conventions and pitfalls**:
  - Handlers are tolerant of multiple signatures: `docx_handler` contains fallbacks for multiple pdf handler API variants. Prefer to adapt handlers rather than changing `app.py`'s dispatch unless the change is repository-wide.
  - `app.py` relies on `detect_type(path: Path)` to map extensions to types — tests rely on exact type strings (`"pdf"`, `"docx"`, `"excel"`, `"image"`).
  - `excel_handler` and `image_handler` import `images_to_pdf` from `pdf_handler.py` via top-level import; ensure `pdf_handler.py` remains importable from project root (sys.path insertion in handlers is present).
  - Outputs are written to `output/<file stem>/...` and combined artifacts often named `annotated_combined.pdf` — tests or integrations may assume these names.

- **When making changes, follow these concrete rules**:
  - If adding or renaming a handler function, also add a compatibility wrapper following the patterns in `docx_handler.py` (e.g., detect new symbol, fall back to expected name).
  - Keep CLI help text and default values in `app.py` in sync with handler option names (e.g., `--resolution`, `--outline`, `--color`, `--font-size`, `--force-render`).
  - Preserve public return shapes (tuples/lists) described above; callers expect `(pages, combined)` or `(images, combined, annotated_xlsx)` forms.

- **Examples to reference in PRs/changes**:
  - Detect type tests: `tests/test_detect_type.py` — keep detect_type behavior stable and case-insensitive.
  - How `app.py` loads pdf handler by path and registers under `pdf_handler` (search for `import_module_by_path` in `app.py`).

If anything above is unclear or you want more examples (e.g., typical test scaffolding or a small local debug script), tell me which area to expand and I will update this file.
