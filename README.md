# openpyxl-simple

A simple wrapper for openpyxl to easily read and write Excel (`.xlsx`) and CSV files.

## Installation

```bash
pip install openpyxl-simple
```

Or using `uv`:

```bash
uv add openpyxl-simple
```

## Quick Start

### Write Data

```python
from datetime import datetime
import openpyxl_simple as ops

# Write list of lists
data = [
    ["Name", "Age", "Website", "Joined"],
    ["Alice", 30, "https://example.com/alice", datetime(2023, 1, 15)],
    ["Bob", 25, "https://example.com/bob", datetime(2024, 6, 1)],
]
ops.write_ll("output.xlsx", data)

# Write list of dicts
header = ["Name", "Age", "Website", "Joined"]
records = [
    {"Name": "Alice", "Age": 30, "Website": "https://example.com/alice", "Joined": datetime(2023, 1, 15)},
    {"Name": "Bob", "Age": 25, "Website": "https://example.com/bob", "Joined": datetime(2024, 6, 1)},
]
ops.write_dict("output.xlsx", header, records)
```

### Read Data

```python
import openpyxl_simple as ops

# Load as list of lists
rows = ops.load_as_list("output.xlsx")
# [
#     ["Name", "Age", "Website", "Joined"],
#     ["Alice", 30, "https://example.com/alice", datetime.datetime(2023, 1, 15, 0, 0)],
#     ["Bob", 25, "https://example.com/bob", datetime.datetime(2024, 6, 1, 0, 0)],
# ]

# Load as list of dictionaries (using the first row as header keys)
records = ops.load_as_dict("output.xlsx")
# [
#     {"Name": "Alice", "Age": 30, "Website": "https://example.com/alice", "Joined": datetime.datetime(2023, 1, 15, 0, 0)},
#     {"Name": "Bob", "Age": 25, "Website": "https://example.com/bob", "Joined": datetime.datetime(2024, 6, 1, 0, 0)},
# ]
```

## Advanced Usage

### Reading Formula vs Calculated Values (`data_only`)

When reading `.xlsx` files containing formulas, pass `data_only=True` via `**kwargs` to get calculated values instead of the formula string:

```python
from openpyxl_simple.reader import load_xlsx_as_dict, load_xlsx_as_list

# Read formulas (default)
formulas = load_xlsx_as_list("data.xlsx", sheet_name="BMI", data_only=False)

# Read calculated cached values
values = load_xlsx_as_dict("data.xlsx", sheet_name="BMI", data_only=True)
```

### Cell Styling (`style_func`)

You can pass a custom styling function to `write_xlsx_ll`, `write_xlsx_dict`, or use the built-in zebra pattern `sample_style_func`:

```python
from openpyxl.styles import PatternFill
from openpyxl_simple.styles import sample_style_func
from openpyxl_simple.writer import write_xlsx_ll

data = [
    ["Name", "Age"],
    ["Alice", 30],
    ["Bob", 25],
]

# Using built-in zebra row striping
write_xlsx_ll("styled.xlsx", data, style_func=sample_style_func)

# Using custom style function
fill_highlight = PatternFill(patternType="solid", fgColor="FFEB3B")

def custom_styler(cell):
    if cell.row == 1:
        cell.fill = fill_highlight

write_xlsx_ll("custom_styled.xlsx", data, style_func=custom_styler)
```

### URLs and Hyperlinks

Strings starting with `http://` or `https://` (without spaces) are automatically converted into clickable Excel hyperlinks when written:

```python
from openpyxl_simple.writer import write_xlsx_dict

header = ["Name", "URL"]
data = [
    {"Name": "Google", "URL": "https://www.google.com"},
    {"Name": "GitHub", "URL": "https://github.com"},
]

write_xlsx_dict("links.xlsx", header, data)
```

### Datetime Handling

`datetime.datetime` and `datetime.date` objects are written directly to Excel cells with native date/time formatting preserved:

```python
from datetime import date, datetime
from openpyxl_simple import load_as_dict, write_dict

header = ["Task", "Due Date", "Created At"]
data = [
    {"Task": "Deploy", "Due Date": date(2026, 9, 10), "Created At": datetime(2026, 9, 5, 12, 0)},
]

write_dict("tasks.xlsx", header, data)
records = load_as_dict("tasks.xlsx")
```

## License

MIT
