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
import openpyxl_simple as oxs

# Load as list of dictionaries (using the first row as header keys)
records = oxs.load_as_dict("sample.xlsx")
# [
#     {"Name": "Alice", "Age": 30, "Website": "https://example.com/alice", "Joined": datetime.datetime(2023, 1, 15, 0, 0)},
#     {"Name": "Bob", "Age": 25, "Website": "https://example.com/bob", "Joined": datetime.datetime(2024, 6, 1, 0, 0)},
# ]

# Load as list of lists
rows = oxs.load_as_list("sample.xlsx")
# [
#     ["Name", "Age", "Website", "Joined"],
#     ["Alice", 30, "https://example.com/alice", datetime.datetime(2023, 1, 15, 0, 0)],
#     ["Bob", 25, "https://example.com/bob", datetime.datetime(2024, 6, 1, 0, 0)],
# ]
```

### Write Data

```python
from datetime import datetime
import openpyxl_simple as oxs

# Write list of dicts
header = ["Name", "Age", "Website", "Joined"]
records = [
    {"Name": "Alice", "Age": 30, "Website": "https://example.com/alice", "Joined": datetime(2023, 1, 15)},
    {"Name": "Bob", "Age": 25, "Website": "https://example.com/bob", "Joined": datetime(2024, 6, 1)},
]
oxs.write_dict("sample.xlsx", header, records)

# Write list of lists
data = [
    ["Name", "Age", "Website", "Joined"],
    ["Alice", 30, "https://example.com/alice", datetime(2023, 1, 15)],
    ["Bob", 25, "https://example.com/bob", datetime(2024, 6, 1)],
]
oxs.write_ll("sample.xlsx", data)
```

## Advanced Usage

For more detailed guides and examples, see [docs/advanced_usage.md](docs/advanced_usage.md):

- **Formulas vs Calculated Values (`data_only`)**: Reading evaluated cached formula results instead of formulas with `data_only=True`.
- **URLs and Hyperlinks**: Automatic hyperlink generation for URLs written to cells.
- **Datetime Handling**: Reading and writing native `date` and `datetime` objects.
- **Cell Styling (`style_func`)**: Applying custom formatting or zebra row striping via styling callbacks.

## License

MIT
