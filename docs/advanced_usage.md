# Advanced Usage

This document covers advanced usage and features of `openpyxl-simple`.

## Reading Formula vs Calculated Values (`data_only`)

When reading `.xlsx` files containing formulas, pass `data_only=True` via `**kwargs` to get calculated values instead of formula strings:

```python
from openpyxl_simple.reader import load_xlsx_as_dict, load_xlsx_as_list

# Read calculated cached values (dict)
values_dict = load_xlsx_as_dict("data.xlsx", sheet_name="BMI", data_only=True)

# Read calculated cached values (list)
values_list = load_xlsx_as_list("data.xlsx", sheet_name="BMI", data_only=True)

# Read formulas (default)
formulas_dict = load_xlsx_as_dict("data.xlsx", sheet_name="BMI", data_only=False)
formulas_list = load_xlsx_as_list("data.xlsx", sheet_name="BMI", data_only=False)
```

## URLs and Hyperlinks

Strings starting with `http://` or `https://` (without spaces) are automatically converted into clickable Excel hyperlinks when written:

```python
from openpyxl_simple.writer import write_xlsx_dict, write_xlsx_ll

# Write dicts with hyperlinks
header = ["Name", "URL"]
records = [
    {"Name": "Google", "URL": "https://www.google.com"},
    {"Name": "GitHub", "URL": "https://github.com"},
]
write_xlsx_dict("links.xlsx", header, records)

# Write list of lists with hyperlinks
data = [
    ["Name", "URL"],
    ["Google", "https://www.google.com"],
    ["GitHub", "https://github.com"],
]
write_xlsx_ll("links_list.xlsx", data)
```

## Datetime Handling

`datetime.datetime` and `datetime.date` objects are written directly to Excel cells with native date/time formatting preserved:

```python
from datetime import date, datetime
from openpyxl_simple import load_as_dict, load_as_list, write_dict, write_ll

# Using dicts
header = ["Task", "Due Date", "Created At"]
records = [
    {"Task": "Deploy", "Due Date": date(2026, 9, 10), "Created At": datetime(2026, 9, 5, 12, 0)},
]
write_dict("tasks.xlsx", header, records)
records_loaded = load_as_dict("tasks.xlsx")

# Using list of lists
data = [
    ["Task", "Due Date", "Created At"],
    ["Deploy", date(2026, 9, 10), datetime(2026, 9, 5, 12, 0)],
]
write_ll("tasks_list.xlsx", data)
rows_loaded = load_as_list("tasks_list.xlsx")
```

## Cell Styling (`style_func`)

You can pass a styling function to `write_xlsx_dict` or `write_xlsx_ll`, or use the built-in zebra pattern `sample_style_func`:

```python
from openpyxl.styles import PatternFill
from openpyxl_simple.styles import sample_style_func
from openpyxl_simple.writer import write_xlsx_dict, write_xlsx_ll

header = ["Name", "Age"]
records = [
    {"Name": "Alice", "Age": 30},
    {"Name": "Bob", "Age": 25},
]
data = [
    ["Name", "Age"],
    ["Alice", 30],
    ["Bob", 25],
]

# Using built-in zebra row striping
write_xlsx_dict("styled_dict.xlsx", header, records, style_func=sample_style_func)
write_xlsx_ll("styled_list.xlsx", data, style_func=sample_style_func)

# Using custom style function
fill_highlight = PatternFill(patternType="solid", fgColor="FFEB3B")

def custom_styler(cell):
    if cell.row == 1:
        cell.fill = fill_highlight

write_xlsx_dict("custom_styled_dict.xlsx", header, records, style_func=custom_styler)
write_xlsx_ll("custom_styled_list.xlsx", data, style_func=custom_styler)
```
