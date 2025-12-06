from pathlib import Path
from app import detect_type

def test_detect_type_pdf():
    assert detect_type(Path("sample.pdf")) == "pdf"

def test_detect_type_docx():
    assert detect_type(Path("file.DOCX")) == "docx"  # case-insensitive

def test_detect_type_excel_xlsx():
    assert detect_type(Path("report.xlsx")) == "excel"

def test_detect_type_excel_xls():
    assert detect_type(Path("report.xls")) == "excel"

def test_detect_type_image_png():
    assert detect_type(Path("img.png")) == "image"

def test_detect_type_invalid():
    assert detect_type(Path("notes.txt")) is None
