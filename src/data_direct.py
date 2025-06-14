__version__ = "1.0.0" 

import traceback
import misc.ui as ui
from venues import venue_report

ui.clear()
print('[Begin Program]')

try:
    venue_report.generate()
except Exception as e:
    print(f'\033[91mAn unexpected error occurred while generating the venue report: {e}')
    traceback.print_exc()
    print('\nThis is a fatal error; the program will now exit. Press any key to continue.')
    ui.pause()