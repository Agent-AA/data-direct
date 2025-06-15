import csv
from datetime import datetime
from dateutil.relativedelta import relativedelta
import misc.ui as ui
import misc.utils as utils
import openpyxl
import os
from venues.records import NoValidSessionsException, VenueRecord

def generate():
    # Display logotype intro
    ui.hideCursor()
    ui.prompt_user('\nThis program will now prompt you to select an Excel (.xlsx) file containing venue data. Press any key to continue.')
    ui.pause()

    # Prompt for excel file
    #file_path = ui.promptFile((('Excel Spreadsheet', ('*.xlsx')),('All files', '*.*')))
    file_path = 'C:\\Users\\alexc\\Documents\\GitHub\\addirectai\\test\\test_input.xlsx'

    # Validate file path
    if file_path == '':
        ui.print_error('No file was selected.')
        ui.pause()
        ui.exit()

    # ----- BEGIN LOADING FILE -----
    try:
        # Read file with openpyxl
        print('Loading Excel file...')
        workboox = openpyxl.load_workbook(file_path, data_only=True)
        sheet = workboox.active
        headers = [cell.value for cell in sheet[1]]

        # Validate headers
        print('Validating file headers...')
        expected_headers = [
            'Job#', 'User', 'MKT', 'LOC#', 'Week', 'Zone', 'Restaurant',
            'St Address', 'City', 'ST', 'ZIP', 'Mail Piece', 'Month', 'Year',
            '# Sessions', 'Qty', 'RSVPs', 'RMI',
            
            'Lunch Day 1', 'Lunch 1 Date', 'Lunch 1 Time',
            'Lunch Day 2', 'Lunch 2 Date', 'Lunch 2 Time',
            'Lunch Day 3', 'Lunch 3 Date', 'Lunch 3 Time',
            'Dinner Day 1', 'Dinner 1 Date', 'Dinner 1 Time',
            'Dinner Day 2', 'Dinner 2 Date', 'Dinner 2 Time',
            'Dinner Day 3', 'Dinner 3 Date', 'Dinner 3 Time',
            ]
        
        missing_headers = [exp_hdr for exp_hdr in expected_headers if exp_hdr not in headers]
        # If there are missing headers
        if len(missing_headers) > 0:
            missing_headers_msg = 'The selected file is missing the following expected columns:'
            for header in missing_headers:
                missing_headers_msg += f'\n{header}'
            ui.print_error(missing_headers_msg)
            ui.pause()
            ui.exit()
    except BaseException as e:
        ui.print_error(f'An error occured while reading the file. This is likely due to invalid file format. See more details below:\n{e}')
        ui.pause()
        ui.exit()
    
    print('Extracting data...')
    venue_records: set[VenueRecord] = set()
    # Load data into structures
    # Iterate through each entry
    for entry in sheet.iter_rows(min_row=2, values_only=True):
        # Convert tuple to dict so we can reference by key
        entry = dict(zip(headers, entry))
        # Create a new venue (or at least try to)
        try:
            new_venue = VenueRecord.from_entry(entry)
            found = False
            # And check if it matches an existing venue
            for existing_venue in venue_records:
                # If there is a match
                if new_venue == existing_venue:
                    # Add new job record to existing venue
                    existing_venue.add_job_record(entry)  
                    found = True

            # If no matching venue found, add this one to the set
            if not found:
                venue_records.add(new_venue)

        except NoValidSessionsException:
            ui.print_warning(f'No valid sessions found for job {entry['Job#']}. Skipping this job.')
        
        except BaseException:
            if entry['Job#'] is not None:
                #ui.print_warning(f'Job {entry['Job#']} is invalidly formatted. Skipping job.')
                # TODO - printing a warning is too verbose. Maybe do something else?
                pass

    ui.print_success('File successfully loaded.')
    print('\nFor default values on any of the following questions, continue without entering.')

    # ----- QUERY USER FOR PARAMETERS -----
    # Query historical data range
    print('Please enter the historical cutoff date for these data.')
    ui.showCursor()
    valid = False
    hist_start_date = None
    while not valid:
        try:
            hist_start_date = utils.parse_datetime(ui.query_user('Start date (MM/DD/YY): ',
                                                  (datetime.now() - relativedelta(months=16)).strftime('%m/%d/%y')))
        except ValueError:
            ui.print_error('The date entered is not valid. Please try again.')
            continue

        valid = True

    # Query scheduling dates
    print('Please enter the scheduling period for this report.')
    valid = False
    start_date, end_date = (None, None)
    while not valid:
        try:
            start_date = utils.parse_datetime(ui.query_user('Start date (MM/DD/YY): '))
            end_date = utils.parse_datetime(ui.query_user('End date (MM/DD/YY): '))
        except ValueError:
            ui.print_error('The entered date is not valid (scheduling dates do not have a default). Please try again.')
            continue

        # Make sure start_date is before end_date
        if start_date > end_date:
            ui.print_error('Start date cannot be after end date. Please try again.')
            continue

        valid = True
    
    # Query minimum RSVPs
    min_rsvps = int(ui.query_user('\nMinimum RSVPs: ', '16'))
    # Query venue cap
    num_venues = int(ui.query_user('Number of venues per market: ', '20'))
    # Query specific markets
    print('\nFor specific markets, use market codes separated by spaces (e.g., "HOU PDX...")')
    markets = ui.query_user('Specific Markets: ').split(' ')

    # SORT DATA
    # Exclude venues...
    # 1. Whose average RSVPs do not meet min_rsvps, and
    # 2. Who have had a job within the last 4 months.
    #
    # Sort venues...
    # If venue had 

    filtered_data = [
        venue for venue in VenueRecord.venue_list if not venue.within_four_months(start_date)
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
    ui.on_error(selected_dir == '',
        "No directory was selected.")

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