import os
import pytest
import shutil
import unittest
from datetime import datetime, timedelta


@pytest.fixture
def tests_dir():
    os.mkdir("_tests")
    yield
    shutil.rmtree("_tests")


class XLSXTest(unittest.TestCase):
    def assertExcelRow(self, sheet, row_num, values, tz=None):
        """
        Asserts the cell values in the given worksheet row. Date values are converted using the provided timezone.
        """
        expected_values = []
        for expected in values:
            # if expected value is datetime, localize and remove microseconds
            if isinstance(expected, datetime):
                expected = expected.astimezone(tz).replace(microsecond=0, tzinfo=None)

            expected_values.append(expected)

        rows = tuple(sheet.rows)

        actual_values = []
        for cell in rows[row_num]:
            actual = cell.value

            if actual is None:
                actual = ""

            if isinstance(actual, datetime):
                actual = actual

            actual_values.append(actual)

        for index, expected in enumerate(expected_values):
            actual = actual_values[index]

            if isinstance(expected, datetime):
                close_enough = abs(expected - actual) < timedelta(seconds=1)
                assert close_enough, f"Datetime value {expected} doesn't match {actual}"
            else:
                assert expected == actual

    def assertExcelSheet(self, sheet, rows, tz=None):
        """
        Asserts the row values in the given worksheet
        """
        assert len(list(sheet.rows)) == len(rows)

        for r, row in enumerate(rows):
            self.assertExcelRow(sheet, r, row, tz)
