__version__ = "1.2.1"


import traceback
import misc.ui as ui
from venues import venue_report

ui.clear(__version__)
print('[Begin Program]')

try:
    venue_report.generate()
except Exception as e:
    ui.print_error(f'An unexpected error occurred while generating the venue report: {e}')
    traceback.print_exc()
    ui.print_error('\nThis is a fatal error; the program will now exit. Press any key to continue.')
    ui.pause()