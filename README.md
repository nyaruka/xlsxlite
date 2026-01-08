# XLSXLite

[![Build Status](https://github.com/nyaruka/xlsxlite/workflows/CI/badge.svg)](https://github.com/nyaruka/xlsxlite/actions?query=workflow%3ACI)
[![Coverage Status](https://codecov.io/gh/nyaruka/xlsxlite/branch/main/graph/badge.svg)](https://codecov.io/gh/nyaruka/xlsxlite)
[![PyPI Release](https://img.shields.io/pypi/v/xlsxlite.svg)](https://pypi.python.org/pypi/xlsxlite/)

This is a lightweight XLSX writer with emphasis on minimizing memory usage. It's also really fast.

```python
from xlsxlite.writer import XLSXBook
book = XLSXBook()
sheet1 = book.add_sheet("People")
sheet1.append_row("Name", "Email", "Age")
sheet1.append_row("Jim", "jim@acme.com", 45)
book.finalize(to_file="simple.xlsx")
```

## Benchmarks

The [benchmarking test](https://github.com/nyaruka/xlsxlite/blob/main/xlsxlite/test/test_perf.py) writes
rows with 10 cells of random string data to a single sheet workbook. The table below gives the times in seconds (lower is better)
to write a spreadsheet with the given number of rows, and includes [xlxswriter](https://xlsxwriter.readthedocs.io/) and
[openpyxl](https://openpyxl.readthedocs.io/) for comparison.

Implementation  | 100,000 rows | 1,000,000 rows
----------------|--------------|---------------
openpyxl        | 43.5         | 469.1
openpyxl + lxml | 21.1         | 226.3
xlsxwriter      | 17.2         | 186.2
xlsxlite        | 1.9          | 19.2

## Limitations

This library is for projects which need to generate large spreadsheets, quickly, for the purposes of data exchange, and
so it intentionally only supports a tiny subset of SpreadsheetML specification:

 * No styling or themes
 * Only strings, numbers, booleans and dates are supported cell types

If you need to do anything fancier then take a look at [xlxswriter](https://xlsxwriter.readthedocs.io/) and
[openpyxl](https://openpyxl.readthedocs.io/).

## Development

To run all tests:

```
uv run pytest xlsxlite -s
```
