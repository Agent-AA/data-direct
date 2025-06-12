__version__ = "1.0.0" 

import csv, ui
import os
import datetime
from venue import Venue

# display logotype intro
ui.clear()
ui.hideCursor()
print("\033[33mThe program will now prompt you to pick a csv file. Press any key to continue.\033[39m")
ui.pause()

# prompt for excel file
file_path = ui.promptFile((('Comma Separated Values File', ('*.csv')),('All files', '*.*')))

# Raise error if file is not a csv file
ui.on_error(not file_path.lower().endswith('.csv'),
    "The selected file is not a csv file.")

with open(file_path, newline='', encoding='utf-8') as f:
    reader = csv.reader(f)
    headers = next(reader)
    data = [dict(zip(headers, row)) for row in reader]
    # Add every venue to the venue_list, but ensure no duplicates
    for entry in data:
        if Venue.not_in_list(Venue(entry)):
            Venue(entry, True)

# Ensure that the loaded file has the correct headers, and if not, tell user.
expected_headers = ['MKT', 'Zone', 'Restaurant', 'St Address', 'City', 'ST', 'ZIP', 'Month', 'Year', 'Lunch 1 Date', 'Lunch 1 Time', 'Lunch 2 Date', 'Lunch 2 Time', 'Lunch 3 Date', 'Lunch 3 Time', 'Lunch 4 Date', 'Lunch 4 Time', 'Dinner 1 Date', 'Dinner 1 Time', 'Dinner 2 Date', 'Dinner 2 Time', 'Dinner 3 Date', 'Dinner 3 Time', 'RSVPs']
missing_headers = [expected_header for expected_header in expected_headers if expected_header not in headers]
# If there are missing headers
if len(missing_headers) > 0:
    missing_headers_msg = "The selected file is missing the following expected columns:"
    for header in missing_headers:
        missing_headers_msg += f'\n{header}'
    # Raise error about missing headers
    ui.raise_error(missing_headers_msg)

# Prompt for date range
ui.clear()
ui.showCursor()

def parse_date(date_str):
    for fmt in ("%m/%d/%y", "%m-%d-%y"):
        try:
            return datetime.datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None

start_date = parse_date(input('Start date (MM-DD-YY): '))
end_date = parse_date(input('\nEnd date (MM-DD-YY): '))

# Validate dates
ui.on_error(start_date is None or end_date is None,
    "Entered an invalid date. Dates must be in the format MM-DD-YY.",
    True)

ui.on_error(end_date < start_date,
    "End date cannot be before start date.",
    True)


filtered_data = [
    venue for venue in Venue.venue_list if not venue.within_four_months(start_date)
]

filtered_data.sort(key=lambda x: (x.attrib('MKT'), -x.average_rsvps))
filtered_data = [
    venue.dict_repr for venue in filtered_data
]

# Write to CSV
ui.clear()
ui.hideCursor()
print("\033[32mReport(s) successfully generated.\033[33m\n\nThe program will now prompt you to pick an output directory. Press any key to continue.\033[39m")
ui.pause()

selected_dir = ui.promptDirectory()
output_dir = selected_dir + f'/VEN_REPORT_{start_date.strftime("%m_%d_%y")}'
os.makedirs(output_dir, exist_ok=True)

from collections import defaultdict
market_groups = defaultdict(list)
for venue in filtered_data:
    market_groups[venue['MKT']].append(venue)

output_files = []
for mkt, venues in market_groups.items():
    filename = f"{mkt}_{start_date.strftime("%m_%d_%y")}.csv"
    file_path = os.path.join(output_dir, filename)
    with open(file_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['MKT','Zone','Restaurant','St Address','City','ST','ZIP'])
        writer.writeheader()
        writer.writerows(venues)
    output_files.append(file_path)

ui.clear()
print(f"\033[32mReport successfully saved to {output_dir}. The program will automatically terminate in 3 seconds.")
ui.wait(3)
ui.exit()