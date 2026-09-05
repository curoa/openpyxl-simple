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

### Read Data

```python
import openpyxl_simple as ops

# Load as list of lists
rows = ops.load_as_list("data.xlsx")

# Load as list of dictionaries (using the first row as header keys)
records = ops.load_as_dict("data.xlsx")
```

### Write Data

```python
import openpyxl_simple as ops

# Write list of lists
data = [
    ["Name", "Age"],
    ["Alice", 30],
    ["Bob", 25],
]
ops.write_ll("output.xlsx", data)

# Write list of dicts
header = ["Name", "Age"]
records = [
    {"Name": "Alice", "Age": 30},
    {"Name": "Bob", "Age": 25},
]
ops.write_dict("output.xlsx", header, records)
```

TODO: use the same as sample data in sample code
TODO: usege about data only, style, url, datetime

## License

MIT
