import datetime


def parse_datetime(datetime_str: str) -> datetime.datetime | None:
    """Parses a datetime string in `MM/YY(YY)`, `MM/DD/YY(YY)`, 
    or `MM/DD/YY(YY) II:MM (AM|PM)` format.
    Returns `None` if the datetime string is formatted invalidly.
    """
    for format in ('%m/%y',             '%m/%Y', 
                   '%m/%d/%y',          '%m/%d/%Y', 
                   '%m/%d/%y %I:%M %p', '%m/%d/%Y %I:%M %p',):
        try:
            return datetime.datetime.strptime(datetime_str, format)
        except ValueError:
            continue
    return None