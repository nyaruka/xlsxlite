from datetime import datetime


def datetime_to_serial(dt):
    """
    Converts the given datetime to the Excel serial format
    """
    if dt.tzinfo:
        raise ValueError("Doesn't support datetimes with timezones")

    temp = datetime(1899, 12, 30)
    delta = dt - temp

    return delta.days + (float(delta.seconds) + float(delta.microseconds) / 1E6) / (60 * 60 * 24)
