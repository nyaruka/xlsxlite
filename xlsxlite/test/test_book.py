import unittest

from xlsxlite.book import XLSXBook


class XLSLiteTest(unittest.TestCase):
    def test_simple_workbook(self):
        book = XLSXBook()
        book.create_sheet("Sheet1")

        book.finalize(to_file="test.xlsx")
