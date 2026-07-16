# OpenPyxl Table

[![PyPI](https://img.shields.io/pypi/v/openpyxl-table)](https://pypi.org/project/openpyxl-table/)
[![Python](https://img.shields.io/pypi/pyversions/openpyxl-table)](https://pypi.org/project/openpyxl-table/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A thin convenience wrapper around [openpyxl](https://openpyxl.readthedocs.io/) that lets you write dictionaries into a worksheet and automatically register the range as a formatted Excel table.

---

## Features

- Write rows from plain `dict` objects — no manual cell addressing required.
- Automatically creates a named, styled Excel table (`ListObject`) on the sheet.
- Column widths can be set inline alongside column names.
- Header row is optional and can be written automatically.
- Fully typed (ships with `py.typed`).

---

## Requirements

- Python ≥ 3.12
- openpyxl ≥ 3.1.5

---

## Installation

```bash
pip install openpyxl-table
```

---

## Quick start

```python
import openpyxl
from openpyxl_table import DictWriter

wb = openpyxl.Workbook()
ws = wb.active

columns = ["Name", "Department", "Score"]

with DictWriter(ws, "A1", columns, auto_header=True) as writer:
    writer.writerow({"Name": "Alice", "Department": "Engineering", "Score": 95})
    writer.writerow({"Name": "Bob",   "Department": "Marketing",   "Score": 87})
    writer.writerow({"Name": "Carol", "Department": "Engineering", "Score": 92})

wb.save("report.xlsx")
```

The resulting workbook contains a table named `Table1` styled with `TableStyleMedium9`.

---

## Usage

### `DictWriter` context manager

```python
DictWriter(
    ws,
    cell,
    fieldnames,
    auto_header=False,
    displayName="Table1",
    style="TableStyleMedium9",
)
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `ws` | `Worksheet` | Target openpyxl worksheet. |
| `cell` | `str \| CellCord` | Top-left cell of the table (e.g. `"B2"` or `CellCord(1, 1)`). |
| `fieldnames` | `list[str \| tuple[str, int]]` | Column names. A `(name, width)` tuple also sets the column width. |
| `auto_header` | `bool` | Write a header row automatically when entering the context. Default: `False`. |
| `displayName` | `str` | Name of the Excel table object. Must be unique in the workbook and must **not** contain spaces. Default: `"Table1"`. |
| `style` | `str` | `TableStyleInfo` name. Default: `"TableStyleMedium9"`. |

Yields a `_DictWriter` instance (see below).

---

### `_DictWriter` methods

#### `writeheader()`

Write the header row using the configured `fieldnames`. Called automatically when `auto_header=True`.

#### `writerow(row)`

Write one data row from a `dict[str, Any]`.  
Keys absent from `fieldnames` are silently ignored; fieldnames absent from `row` leave the cell empty.

#### `writerows(rows)`

Write multiple data rows at once (calls `writerow` for each item).

```python
data = [
    {"Name": "Alice", "Score": 95},
    {"Name": "Bob",   "Score": 87},
]

with DictWriter(ws, "A1", ["Name", "Score"], auto_header=True) as writer:
    writer.writerows(data)
```

---

### Column widths

Pass `(column_name, width)` tuples in `fieldnames` to set column widths:

```python
columns = [
    ("Name",  25),   # 25 character units wide
    ("Score", 10),
]

with DictWriter(ws, "A1", columns, auto_header=True) as writer:
    writer.writerow({"Name": "Alice", "Score": 95})
```

You can mix plain strings and tuples freely:

```python
columns = ["ID", ("Name", 30), "Score"]
```

---

### Multiple tables on one sheet

Use `displayName` to give each table a unique name:

```python
with DictWriter(ws, "A1", ["Name", "Score"], auto_header=True, displayName="Scores") as w:
    w.writerows(scores)

with DictWriter(ws, "D1", ["City", "Country"], auto_header=True, displayName="Locations") as w:
    w.writerows(locations)
```

---

### `CellCord`

A zero-based `(column, row)` coordinate used internally (and available publicly for advanced use):

```python
from openpyxl_table import CellCord

c = CellCord.from_string("B3")  # → CellCord(column=1, row=2)
c.offset(column=2, row=1)       # → CellCord(column=3, row=3)
```

---

## License

MIT – see [LICENSE](LICENSE).
