from collections import defaultdict
from datetime import datetime
from dateutil.relativedelta import relativedelta
import misc.ui as ui
import misc.utils as utils
import openpyxl
import os
from tqdm import tqdm
from venues.records import NoValidSessionsException, VenueRecord

def generate():
    # Display logotype intro
    ui.hideCursor()
    ui.prompt_user('\nThis program will now prompt you to select an Excel (.xlsx) file containing venue data. Press any key to continue.')
    ui.pause()

    # Prompt for excel file
    file_path = ui.promptFile((('Excel Spreadsheet', ('*.xlsx')),('All files', '*.*')))
    #file_path = 'C:\\Users\\alexc\\Documents\\GitHub\\addirectai\\test\\test_input.xlsx'

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
    
    # Query for historical data range
    print('Please enter the historical cutoff date for these data (default is 16 months prior to today).')
    ui.showCursor()
    valid = False
    cutoff_date = None
    while not valid:
        try:
            cutoff_date = utils.parse_datetime(ui.query_user('Start date (MM/DD/YY): ',
                                                  (datetime.now() - relativedelta(months=16)).strftime('%m/%d/%y')))
        except ValueError:
            ui.print_error('The date entered is not valid. Please try again.')
            continue

        valid = True

    print('Extracting data. This may take a minute...')
    venue_records: set[VenueRecord] = set()
    # Load data into structures
    # Iterate through each entry
    for entry in tqdm(sheet.iter_rows(min_row=2, values_only=True), total=sheet.max_row - 1):
        # If entry contains a date before cutoff date, don't evaluate
        outdated = False
        for val in entry:
            try:
                if val < cutoff_date:
                    outdated = True
                    break
            except:
                pass
        if outdated:
            continue

        # Convert tuple to dict so we can reference by key
        entry = dict(zip(headers, entry))

        # If no job id, then there's not really an entry here
        if entry['Job#'] is None:
            continue

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
            #tqdm.write(ui.warning(f'No valid sessions found for job {entry['Job#']}. Skipping this job.'))
            pass

        except (TypeError, ValueError):
            #if entry['Job#'] is not None:
            #    tqdm.write(ui.warning(f'Job {entry['Job#']} is invalidly formatted. Skipping this job.'))
            pass

        #except BaseException:
        #    if entry['Job#'] is not None:
        #        #ui.print_warning(f'Job {entry['Job#']} is invalidly formatted. Skipping job.')
        #        # TODO - printing a warning is too verbose. Maybe do something else?
        #        pass

    ui.print_success('Extraction complete.')

    # ----- QUERY USER FOR PARAMETERS -----
    # Query scheduling dates
    print('\nPlease enter the scheduling period for this report. Scheduling period does not support default values.')
    valid = False
    start_date, end_date = (None, None)
    while not valid:
        try:
            start_date = utils.parse_datetime(ui.query_user('Start date (MM/DD/YY): '))
            end_date = utils.parse_datetime(ui.query_user('End date (MM/DD/YY): '))
        except ValueError:
            ui.print_error('The entered date is not valid. Please try again.')
            continue

        # Make sure start_date is before end_date
        if start_date > end_date:
            ui.print_error('Start date cannot be after end date. Please try again.')
            continue

        valid = True
    
    
    print('\nFor default values on any of the following questions, continue without entering anything.')
    # Query minimum RSVPs
    min_rsvps = int(ui.query_user('Minimum RSVPs: ', '16'))
    # Query venue cap
    num_venues = int(ui.query_user('Number of venues per market: ', '20'))
    # Query specific markets
    print('\nFor specific markets, use market codes separated by spaces (e.g., "HOU PDX...")')
    markets = ui.query_user('Specific Markets: ').split(' ')

    # SORT DATA
    # Exclude venues...
    # 1. Whose average RSVPs do not meet min_rsvps, and
    # 2. Who are in a zone which has had a job within the last four months
    #
    # Sort venues...
    # If venue had a session within two weeks last year of scheduling period 

    print('Executing set exclusions...')
    # We want to exclude all zones that have had an event within four months
    saturated_zones = {
        venue.zone for venue in venue_records if venue.within_four_months(start_date)
    }

    # Filter by saturated zones and minimum rsvps
    filtered_data = {
        venue for venue in venue_records
        if (venue.zone not in saturated_zones
            and venue.average_rsvps >= min_rsvps)
    }

    print('Performing optimizations...')
    sorted_by_rsvps = sorted(filtered_data, key=lambda venue: venue.average_rsvps, reverse=True)
    sorted_data = sorted(sorted_by_rsvps, key=lambda venue: venue.around_time_last_year(start_date, end_date, prox_weeks=2), reverse=True)

    ui.print_success('Exclusions and optimizations complete.')

    # Prepare to output data
    ui.prompt_user('\nThis program will now prompt you to select an ouput directory. Press any key to continue.')
    ui.pause()
    
    selected_dir = ui.promptDirectory()
    #selected_dir = 'C:\\Users\\alexc\\Documents\\GitHub\\addirectai\\test'

    if selected_dir == '':
        ui.print_warning('No directory selected. Terminating program.')
        ui.pause()
        ui.exit()

    print('Creating output directory...')
    output_dir = selected_dir + f'\\VEN_REPORT_{start_date.strftime("%m_%d_%y")}-{end_date.strftime("%m_%d_%y")}'
    os.makedirs(output_dir, exist_ok=True)

    print('Classifying records by market...')
    venues_by_market = defaultdict(list[VenueRecord])
    for venue in sorted_data:
        venues_by_market[venue.market].append(venue)
    
    print('Writing records to new files...')
    # Write to new excel file
    for market, venues in venues_by_market.items():
        # If user requested specific markets, halt for non-specified markets
        if markets[0] != '' and market not in markets:
            continue

        wb = openpyxl.Workbook()
        ws = wb.active

        headers = [
            'Job#', 'User', 'MKT', 'LOC#', 'Week', 'Zone',
            'Restaurant', 'St Address', 'City', 'ST', 'ZIP',
            'Mail Piece', 'Month', 'Year', '# Sessions',
            'Session Type', 'Qty', 'RSVPs', 'Average RSVPs', 'RMI']

        ws.append(headers)

        # This is for capping number of written venues
        i = 0

        for venue in venues:
            if i < num_venues:
                row = venue.to_entry()
                ws.append(row)
                i += 1

        file_path = os.path.join(output_dir, f'{market}_{start_date.strftime("%m_%d_%y")}-{end_date.strftime("%m_%d_%y")}.xlsx')
        wb.save(file_path)

    ui.print_success(f"Report(s) have been saved. You can safely close the program.")
    ui.pause()
    ui.exit()