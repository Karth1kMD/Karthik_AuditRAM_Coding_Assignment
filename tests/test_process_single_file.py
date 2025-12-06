from pathlib import Path
import builtins
import app

def fake_pdf_annotate(path, query, out_dir, resolution, outline_width, stroke_color, prefer_fallback, force_render):
    # Simulate page list and combined path (string or Path)
    return (["page1.png", "page2.png"], Path(out_dir) / "combined.pdf")

def fake_docx_process(path, query, out_dir, resolution, outline_width, stroke_color, force_render, prefer_fallback=True):
    return (["doc_page.png"], Path(out_dir) / "combined_docx.pdf")

def fake_excel_process(path, query, out_dir, stroke_color, outline_width, font_size):
    return (["sheet1.png"], "combined_excel.pdf", Path(out_dir) / "annotated.xlsx")

def fake_image_process(path, query, out_dir, stroke_color, outline_width):
    return (str(Path(out_dir) / "img.png"), None)

def test_process_single_file_pdf(monkeypatch, tmp_path):
    # Create fake pdf file
    p = tmp_path / "file.pdf"
    p.write_text("x")
    outdir = tmp_path / "out"

    # Ensure app.pdf_handler exists and monkeypatch its function
    class PDF:
        pass
    PDF.annotate_pdf_and_build_combined = staticmethod(fake_pdf_annotate)
    monkeypatch.setattr(app, "pdf_handler", PDF)

    result, subdir = app.process_single_file(p, "kw", outdir, {
        "resolution": 150, "outline": 3, "color": "red", "force_render": False
    })

    assert "pages" in result
    assert isinstance(result["pages"], list)
    assert result["combined"] is not None
    assert subdir.exists()

def test_process_single_file_docx(monkeypatch, tmp_path):
    p = tmp_path / "file.docx"
    p.write_text("x")
    outdir = tmp_path / "out2"

    class DOCX:
        pass
    DOCX.process_word_file = staticmethod(fake_docx_process)
    monkeypatch.setattr(app, "docx_handler", DOCX)

    result, subdir = app.process_single_file(p, "kw", outdir, {
        "resolution": 150, "outline": 2, "color": "blue", "force_render": False
    })

    assert result["combined"] is not None
    assert subdir.exists()

def test_process_single_file_excel(monkeypatch, tmp_path):
    p = tmp_path / "file.xlsx"
    p.write_text("x")
    outdir = tmp_path / "out3"

    class EX:
        pass
    EX.process_excel_file = staticmethod(fake_excel_process)
    monkeypatch.setattr(app, "excel_handler", EX)

    result, subdir = app.process_single_file(p, "kw", outdir, {
        "outline": 2, "color": "green", "font_size": 12
    })

    assert "annotated_xlsx" in result
    assert subdir.exists()

def test_process_single_file_image(monkeypatch, tmp_path):
    p = tmp_path / "img.png"
    p.write_text("x")
    outdir = tmp_path / "out4"

    class IM:
        pass
    IM.process_image = staticmethod(fake_image_process)
    monkeypatch.setattr(app, "image_handler", IM)

    result, subdir = app.process_single_file(p, "kw", outdir, {
        "outline": 1, "color": "red"
    })

    assert "pages" in result
    assert subdir.exists()
