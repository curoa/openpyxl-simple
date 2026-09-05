import csv
import os

import openpyxl

from openpyxl_simple.utils import guess_utf_encoding, to_path

def load_as_list(fpath):
    fpath = to_path(fpath)
    if fpath.suffix == ".xlsx" and fpath.exists():
        return load_xlsx_as_list(fpath)
    else:
        fpath = os.path.splitext(fpath)[0] + ".csv"
        return load_csv_as_list(fpath)

# use `data_only=True` to get calculated value
def load_xlsx_as_list(fname, sheet_name=None, **kwargs):
    wb = openpyxl.load_workbook(fname, **kwargs)
    #print('wb.sheetnames', wb.sheetnames) # debug
    if sheet_name is None:
        ws = wb.active
    else:
        ws = wb[sheet_name]
    return load_ws_as_list(ws)

# use `data_only=True` to get calculated value
def load_ws_as_list(ws):
    data = []
    for row in ws.iter_rows():
        one = []
        for cell in row:
            one.append(cell.value)
        data.append(one)
    return data

def load_csv_as_list(fpath, encoding="utf-16"):
    return list(csv.reader(open(fpath, encoding=encoding)))

def load_as_dict(fpath):
    fpath = to_path(fpath)
    if fpath.suffix == ".xlsx" and fpath.exists():
        return load_xlsx_as_dict(fpath)
    else:
        fpath = os.path.splitext(fpath)[0] + ".csv"
        return load_csv_as_dict(fpath)

# use `data_only=True` to get calculated value
def load_xlsx_as_dict(fpath, sheet_name=None, **kwargs): #TODO rename
    wb = openpyxl.load_workbook(fpath, **kwargs)
    if sheet_name is None:
        ws = wb.active
    else:
        ws = wb[sheet_name]
    return load_ws_as_dict(ws)

def load_ws_as_dict(ws):
    header = []
    for row in ws.iter_rows(min_row=1):
        for cell in row:
            header.append(cell.value)
        break
    data = []
    for row in ws.iter_rows(min_row=2):
        one = {}
        for i, cell in enumerate(row):
            one[header[i]] = cell.value
        data.append(one)
    return data

def load_csv_as_dict(fpath, encoding=None):
    if encoding is None:
        encoding = guess_utf_encoding(fpath)
    reader = csv.DictReader(open(fpath, encoding=encoding))
    data = [row for row in reader]
    return data
