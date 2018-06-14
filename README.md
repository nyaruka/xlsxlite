# XLSXLite

[![Build Status](https://travis-ci.org/nyaruka/xlsxlite.svg?branch=master)](https://travis-ci.org/nyaruka/xlsxlite)
[![Coverage Status](https://coveralls.io/repos/github/nyaruka/xlsxlite/badge.svg?branch=master)](https://coveralls.io/github/nyaruka/xlsxlite?branch=master)
[![PyPI Release](https://img.shields.io/pypi/v/xlsxlite.svg)](https://pypi.python.org/pypi/xlsxlite/)

This is a lightweight XLSX writer with emphasis on minimizing memory usage. It's also really fast.

```python
from xlsxlite.book import XLSXBook
book = XLSXBook()
sheet1 = book.add_sheet("People")
sheet1.append_row("Name", "Email", "Age")
sheet1.append_row("Jim", "jim@acme.com", 45)
book.finalize(to_file="simple.xlsx")
```

## Benchmarks

The [benchmarking test](https://github.com/nyaruka/xlsxlite/blob/master/xlsxlite/test/test_perf.py) writes
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
 * Only strings, numbers and dates are supported cell types

If you need to do anything fancier then take a look at [xlxswriter](https://xlsxwriter.readthedocs.io/) and
[openpyxl](https://openpyxl.readthedocs.io/).

## Development

To run all tests:

```
py.test xlsxlite -s
```
