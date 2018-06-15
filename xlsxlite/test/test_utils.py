import pytest
import pytz

from datetime import datetime
from xlsxlite.utils import datetime_to_serial


def test_datetime_to_serial():
    assert datetime_to_serial(datetime(2013, 1, 1, 12, 0, 0)) == 41275.5
    assert datetime_to_serial(datetime(2018, 6, 15, 11, 24, 30, 0)) == 43266.47534722222

    # try with a non-naive datetime
    with pytest.raises(ValueError):
        datetime_to_serial(datetime(2018, 6, 15, 11, 24, 30, 0, pytz.UTC))
