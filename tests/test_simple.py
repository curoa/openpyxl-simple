from pathlib import Path
import openpyxl_simple as oxs
from openpyxl_simple.decorator import deco_fname_check, deco_to_path
from openpyxl_simple.encoding import guess_utf_encoding
from openpyxl_simple.table import collect_header, tie_key_value


def test_exports():
    for name in oxs.__all__:
        assert hasattr(oxs, name), f"Missing {name}"


def test_deco_to_path():
    @deco_to_path
    def dummy_fpath(fpath):
        return fpath

    assert isinstance(dummy_fpath("foo.txt"), type(Path("foo.txt")))
    assert isinstance(dummy_fpath(fpath="foo.txt"), type(Path("foo.txt")))
    assert isinstance(dummy_fpath(Path("foo.txt")), type(Path("foo.txt")))

    @deco_to_path(arg_name="filepath")
    def dummy_filepath(filepath):
        return filepath

    assert isinstance(dummy_filepath("bar.txt"), type(Path("bar.txt")))
    assert isinstance(dummy_filepath(filepath="bar.txt"), type(Path("bar.txt")))

    @deco_to_path
    def dummy_default(fpath="default.txt"):
        return fpath

    assert isinstance(dummy_default(), type(Path("default.txt")))
    assert dummy_default() == Path("default.txt")


def test_deco_fname_check():
    @deco_fname_check("xlsx")
    def dummy_save(fpath):
        return True

    assert dummy_save("test.xlsx") is True
    assert dummy_save("test.csv") is True


def test_encoding(tmp_path):
    f = tmp_path / "test_utf8.txt"
    f.write_text("hello", encoding="utf-8")
    assert guess_utf_encoding(str(f)) == "utf-8"

    f_bom = tmp_path / "test_bom.txt"
    f_bom.write_bytes(b"\xef\xbb\xbfhello")
    assert guess_utf_encoding(str(f_bom)) == "utf-8-sig"


def test_table():
    header = ["id", "name"]
    rows = [[1, "Alice"], [2, "Bob"]]
    data = tie_key_value(header, rows)
    assert data == [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}]
    collected = collect_header(data)
    assert collected == ["id", "name"]
