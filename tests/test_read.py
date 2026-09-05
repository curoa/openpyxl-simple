import csv
from pathlib import Path
import openpyxl
from openpyxl_simple.reader import (
    load_as_dict,
    load_as_list,
    load_csv_as_dict,
    load_csv_as_list,
    load_ws_as_dict,
    load_ws_as_list,
    load_xlsx_as_dict,
    load_xlsx_as_list,
)

TEST_DIR = Path(__file__).parent / "data"
SAMPLE_XLSX = TEST_DIR / "sample.xlsx"
SAMPLE_CSV = TEST_DIR / "sample.csv"


def test_read_ws_as_list():
    wb = openpyxl.load_workbook(SAMPLE_XLSX)
    ws = wb.active
    result = load_ws_as_list(ws)
    assert result == [["number"], ["AE3803"], ["AA5100"], ["NT4201"]]


def test_read_ws_as_dict():
    wb = openpyxl.load_workbook(SAMPLE_XLSX)
    ws = wb.active
    result = load_ws_as_dict(ws)
    assert result == [
        {"number": "AE3803"},
        {"number": "AA5100"},
        {"number": "NT4201"},
    ]


def test_read_xlsx():
    assert load_xlsx_as_list(str(SAMPLE_XLSX)) == [["number"], ["AE3803"], ["AA5100"], ["NT4201"]]
    assert load_xlsx_as_dict(str(SAMPLE_XLSX)) == [
        {"number": "AE3803"},
        {"number": "AA5100"},
        {"number": "NT4201"},
    ]
    assert load_xlsx_as_list(str(SAMPLE_XLSX), sheet_name="アニメ") == [
        ["はたらく細胞", "2018年7月8日"],
        ["化物語", "2009年7月3日"],
    ]


def test_read_xlsx_data_only():
    list_formula = load_xlsx_as_list(str(SAMPLE_XLSX), sheet_name="BMI", data_only=False)
    assert list_formula == [
        ["身長(cm)", "体重(kg)", "BMI"],
        [158, 45, "=B2/((A2/100)^2)"],
        [153, 38, "=B3/((A3/100)^2)"],
    ]

    dict_formula = load_xlsx_as_dict(str(SAMPLE_XLSX), sheet_name="BMI", data_only=False)
    assert dict_formula == [
        {"身長(cm)": 158, "体重(kg)": 45, "BMI": "=B2/((A2/100)^2)"},
        {"身長(cm)": 153, "体重(kg)": 38, "BMI": "=B3/((A3/100)^2)"},
    ]

    list_val = load_xlsx_as_list(str(SAMPLE_XLSX), sheet_name="BMI", data_only=True)
    assert list_val[0] == ["身長(cm)", "体重(kg)", "BMI"]
    assert list_val[1][0] == 158 and list_val[1][1] == 45
    assert round(list_val[1][2], 2) == 18.03
    assert list_val[2][0] == 153 and list_val[2][1] == 38
    assert round(list_val[2][2], 2) == 16.23

    dict_val = load_xlsx_as_dict(str(SAMPLE_XLSX), sheet_name="BMI", data_only=True)
    assert dict_val[0]["身長(cm)"] == 158 and dict_val[0]["体重(kg)"] == 45
    assert round(dict_val[0]["BMI"], 2) == 18.03
    assert dict_val[1]["身長(cm)"] == 153 and dict_val[1]["体重(kg)"] == 38
    assert round(dict_val[1]["BMI"], 2) == 16.23


def test_read_csv():
    assert load_csv_as_list(str(SAMPLE_CSV), encoding="utf-16") == [
        ["はたらく細胞", "2018年7月8日"],
        ["化物語", "2009年7月3日"],
    ]


def test_read_dispatcher_fallback(tmp_path):
    csv_file = tmp_path / "fallback.csv"
    with open(csv_file, "w", encoding="utf-16", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["c1", "c2"])
        writer.writerow(["v1", "v2"])

    non_exist_xlsx = tmp_path / "fallback.xlsx"
    assert load_as_list(str(non_exist_xlsx)) == [["c1", "c2"], ["v1", "v2"]]

    csv_dict_file = tmp_path / "fallback_dict.csv"
    with open(csv_dict_file, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["key", "value"])
        writer.writeheader()
        writer.writerow({"key": "val1", "value": "val2"})

    non_exist_xlsx_dict = tmp_path / "fallback_dict.xlsx"
    assert load_as_dict(str(non_exist_xlsx_dict)) == [{"key": "val1", "value": "val2"}]
