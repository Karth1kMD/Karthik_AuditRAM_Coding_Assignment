from pathlib import Path
from app import gather_files, detect_type

def test_gather_files_single_file(tmp_path):
    # create a fake file with a supported extension
    f = tmp_path / "doc1.pdf"
    f.write_text("dummy")
    result = gather_files(f, recursive=False)
    assert isinstance(result, list)
    assert len(result) == 1
    assert result[0].name == "doc1.pdf"

def test_gather_files_directory_non_recursive(tmp_path):
    # create files and a subfolder
    base = tmp_path / "root"
    base.mkdir()
    (base / "a.pdf").write_text("a")
    (base / "b.txt").write_text("b")  # unsupported
    sub = base / "sub"
    sub.mkdir()
    (sub / "c.pdf").write_text("c")

    # non-recursive should return only a.pdf
    files = gather_files(base, recursive=False)
    names = sorted([p.name for p in files])
    assert names == ["a.pdf"]

def test_gather_files_directory_recursive(tmp_path):
    base = tmp_path / "root2"
    base.mkdir()
    (base / "a.pdf").write_text("a")
    sub = base / "sub2"
    sub.mkdir()
    (sub / "c.pdf").write_text("c")
    files = gather_files(base, recursive=True)
    names = sorted([p.name for p in files])
    assert names == ["a.pdf", "c.pdf"]
