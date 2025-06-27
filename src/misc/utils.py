import datetime


def parse_datetime(datetime_str: str) -> datetime.datetime:
    """Parses a datetime string in `MM/YY(YY)`, `MM/DD/YY(YY)`, 
    or `MM/DD/YY(YY) II:MM (AM|PM)` format.
    Raises a ValueError if the datetime string is formatted invalidly.
    """
    for format in ('%m/%y',             '%m/%Y', 
                   '%m/%d/%y',          '%m/%d/%Y', 
                   '%m/%d/%y %I:%M %p', '%m/%d/%Y %I:%M %p',):
        try:
            return datetime.datetime.strptime(datetime_str, format)
        except ValueError:
            continue
    raise ValueError(f'Datetime string {datetime_str} is not in a valid format.')

def parse_month_year(month: str, year: str) -> datetime.datetime:
    """
    Parses a month and year string into a datetime object representing the first day of that month.
    Args:
        month (str): Month as a string (e.g., '1', '01', 'Jan', 'January').
        year (str): Year as a string (e.g., '2024', '24').
    Returns:
        datetime.datetime: Datetime object for the first day of the given month and year.
    Raises:
        ValueError: If the month or year cannot be parsed.
    """
    try:
        # Try parsing month as integer
        try:
            month_int = int(month)
        except ValueError:
            # Try parsing month as name
            month_int = datetime.datetime.strptime(month[:3], "%b").month
        # Handle 2-digit years
        year_int = int(year)
        if year_int < 100:
            year_int += 2000 if year_int < 70 else 1900
        return datetime.datetime(year_int, month_int, 1)
    except Exception as e:
        raise ValueError(f"Invalid month/year: {month}/{year}") from e
