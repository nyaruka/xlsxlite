# Introduction

[![Build Status](https://travis-ci.org/nyaruka/xlsxlite.svg?branch=master)](https://travis-ci.org/nyaruka/xlsxlite)
[![Coverage Status](https://coveralls.io/repos/github/nyaruka/xlsxlite/badge.svg?branch=master)](https://coveralls.io/github/nyaruka/xlsxlite?branch=master)
[![PyPI Release](https://img.shields.io/pypi/v/xlsxlite.svg)](https://pypi.python.org/pypi/xlsxlite/)

XLSXLite is a lightweight XLSX writer with emphasis on minimizing memory usage.

```python
from xlsxlite.book import XLSXBook
book = XLSXBook()
sheet1 = book.add_sheet("People")
sheet1.append_row("Name", "Email")
sheet1.append_row("Jim", "jim@acme.com")
book.finalize(to_file="simple.xlsx")
```

## Development

To run all tests:

```
py.test xlsxlite -s
```
