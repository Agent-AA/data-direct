from collections import defaultdict
from datetime import datetime
from typing import overload
from dateutil.relativedelta import relativedelta
import misc.ui as ui
import misc.utils as utils
import openpyxl
import os
from tqdm import tqdm
from venues.records import VenueRecord
from venues.errors import HashError, NoValidSessionsException

expected_headers = [
            'Job#', 'User', 'MKT', 'LOC#', 'Week', 'Zone', 'Restaurant',
            'St Address', 'City', 'ST', 'ZIP', 'Mail Piece', 'Month', 'Year',
            '# Sessions', 'Qty', 'RSVPs', 'RMI',
            
            'Lunch Day 1', 'Lunch 1 Date', 'Lunch 1 Time',
            'Lunch Day 2', 'Lunch 2 Date', 'Lunch 2 Time',
            'Lunch Day 3', 'Lunch 3 Date', 'Lunch 3 Time',
            'Dinner Day 1', 'Dinner 1 Date', 'Dinner 1 Time',
            'Dinner Day 2', 'Dinner 2 Date', 'Dinner 2 Time',
            'Dinner Day 3', 'Dinner 3 Date', 'Dinner 3 Time']

@overload
def generate() -> None: ...
@overload
def generate(venue_records: set['VenueRecord']) -> None: ...

def generate(venue_records: set['VenueRecord']=None):
    print('\n[Begin new report]')
    if (venue_records is None):
        # Display logotype intro
        ui.hideCursor()
        ui.prompt_user('\nThis program will now prompt you to select an Excel (.xlsx) file containing venue data. Press any key to continue.')
        ui.pause()

        # Prompt for excel file
        file_path = _get_file_path()

        # Load file
        headers, raw_data_sheet = _load_excel(file_path)
        
        # Query for historical data range
        print('Please enter the historical cutoff date for these data (default is 16 months prior to today).')
        cutoff_date = ui.query_date(
            'Cutoff date (MM/DD/YY): ',
            default=datetime.now() - relativedelta(months=16))
        
        print('Extracting data. This may take a minute...')
        venue_records = _extract_data(headers, raw_data_sheet, cutoff_date)
        ui.print_success('Extraction complete.')


    # ----- QUERY USER FOR PARAMETERS -----
    # Query scheduling dates
    print('\nPlease enter the scheduling period for this report. Scheduling period does not support default values.')
    
    start_date = ui.query_date('Start date (MM/DD/YY): ')
    end_date = ui.query_date('End Date (MM/DD/YY): ')

    # We can't have the start date be after the end date.
    while start_date > end_date:
        ui.print_error('Scheduling period start date cannot be after end date. Please try again.')
        start_date = ui.query_date('Start date (MM/DD/YY): ')
        end_date = ui.query_date('End Date (MM/DD/YY): ')

    
    # A bunch of other queries
    print('\nFor default values on any of the following questions, continue without entering anything.')
    # Query saturation period
    saturation_period = int(ui.query_user('Zone Saturation Period (weeks): ', '16'))
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
    filtered_data = _filter_data(venue_records, saturation_period, start_date, min_rsvps)

    # Rank venues by latest job ROR 
    # and then categorize by whether there was a job
    # around the same time last year
    print('Performing optimizations...')
    sorted_by_rsvps = sorted(filtered_data, key=lambda venue: venue.latest_job.ror, reverse=True)
    sorted_data = sorted(sorted_by_rsvps, key=lambda venue: venue.around_time_last_year(start_date, end_date, prox_weeks=2), reverse=True)

    ui.print_success('Exclusions and optimizations complete.')

    # Prepare to output data
    ui.prompt_user('\nThis program will now prompt you to select an ouput directory. Press any key to continue.')
    ui.pause()
    
    selected_dir = ui.promptDirectory()
    #selected_dir = 'C:\\Users\\alexc\\Documents\\GitHub\\addirectai\\test'

    if selected_dir == '':
        ui.print_error('No directory selected. Terminating report.')
        generate(venue_records)
        return

    print('Creating output directory...')
    output_dir = selected_dir + f'\\VEN_REPORT_{start_date.strftime("%m_%d_%y")}-{end_date.strftime("%m_%d_%y")}'

    # Check if directory already exists, and if so warn user. 
    try:
        os.makedirs(output_dir, exist_ok=False)
    except OSError:
        ui.print_warning('WARNING: A folder already exists at the selected location and will be overwritten. Press N to abort. Press any other key to continue.')
        keypress = ui.pause()

        if (keypress == b'n' or keypress == b'N'):
            print('Terminating report.')
            generate(venue_records)
            return
        
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

        # IMPORTANT if headers are changed, then the returned tuple
        # from Venue.to_entry() must also be changed so that the header
        # order matches the data order.
        headers = [
            'Job#', 'User', 'MKT', 'LOC#', 'Week', 'Zone',
            'Restaurant', 'St Address', 'City', 'ST', 'ZIP',
            'Mail Piece', 'Month', 'Year', '# Sessions',
            'Session Type', 'Qty', 'RSVPs', 'RMI', 'ROR (%)', 
            'Average RSVPs', 'Average ROR (%)']

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

    ui.print_success(f"Report(s) have been saved. Press any key to begin a new report, or close the program.")
    ui.pause()
    
    generate(venue_records)
    return



# ===== internal helper functions ===== #



def _get_file_path(test: bool=False) -> str:
    """Query the user for a filepath and returns a string. Will cause a UI
    error and exit if the user enters an empty string.
    """
    if test:
        file_path = 'C:\\Users\\alexc\\Documents\\GitHub\\addirectai\\test\\test_input.xlsx'
    else:
        file_path = ui.promptFile((('Excel Spreadsheet', ('*.xlsx')),('All files', '*.*')))

    # Validate file path
    if file_path == '':
        ui.print_error('No file was selected.')
        ui.pause()
        ui.exit()
    
    return file_path


def _load_excel(excel_file_path: str) -> list:
    """Attempt to load an excel spreadsheet and return a list which
    can be iterated over. Will cause a UI error if an exception occurs or
    the file is missing necessary file headers, which are hardcoded in this
    class.
    """
    try:
        # Read file with openpyxl
        print('Loading Excel file...')
        workbook = openpyxl.load_workbook(excel_file_path, data_only=True)
        sheet = workbook.active
        headers = [cell.value for cell in sheet[1]]
        
        missing_headers = [exp_hdr for exp_hdr in expected_headers if exp_hdr not in headers]
        # If there are missing headers
        if len(missing_headers) > 0:
            missing_headers_msg = 'The selected file is missing the following expected columns:'
            for header in missing_headers:
                missing_headers_msg += f'\n{header}'
            ui.print_error(missing_headers_msg)
            ui.pause()
            ui.exit()
        else:
            ui.print_success('All expected headers are present.')

    except BaseException as e:
        ui.print_error(f'An error occured while reading the file. This is likely due to invalid file format.')
        ui.pause()
        ui.exit()
    
    return (headers, sheet)


def _extract_data(headers: list[str], raw_data_sheet: list, cutoff_date: datetime) -> set['VenueRecord']:
    """Accepts a data sheet like one from an openpyxl workbook and retrns
    a set of VenueRecords. Will skip over malformed entries in the sheet
    without raising any exceptions.
    """
    venue_records: set[VenueRecord] = set()
    # Load data into structures
    # Iterate through each entry
    for entry in tqdm(raw_data_sheet.iter_rows(min_row=2, values_only=True), total=raw_data_sheet.max_row - 1):
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

        except (HashError):
            pass

        #except BaseException:
        #    if entry['Job#'] is not None:
        #        #ui.print_warning(f'Job {entry['Job#']} is invalidly formatted. Skipping job.')
        #        # TODO - printing a warning is too verbose. Maybe do something else?
        #        pass

    return venue_records


def _filter_data(venue_records: set[VenueRecord], saturation_period: int, start_date: datetime, min_rsvps: int):
    """Filters out undesirable venues. The current criteria is based on minimum
    number of RSVPs and whether a venue's zone has had a seminar within the `saturation_period` (weeks).
    """
    
    saturated_zones = {
        venue.zone for venue in venue_records 
        if venue.within(weeks=saturation_period, ref_date=start_date)
    }

    # Filter by saturated zones and minimum rsvps
    filtered_data = {
        venue for venue in venue_records
        if (venue.zone not in saturated_zones
            and venue.average_rsvps >= min_rsvps)
    }

    return filtered_data