import os, sys, msvcrt, time
import tkinter
import tkinter.filedialog
from datetime import datetime
from misc import utils, ui


def print_error(msg: str):
    """Prints a message in bright red text to the terminal.
    """
    print(f"\033[91m{msg}\033[39m")

def print_warning(msg: str):
    """Prints a message in bright yellow text to the terminal.
    """
    print(warning(msg))

def warning(msg: str) -> str:
    return f"\033[93m{msg}\033[39m"

def print_success(msg: str):
    """Prints a message in bright green text to the terminal.
    """
    print(f"\033[92m{msg}\033[39m")

def prompt_user(msg: str):
    """Prints a message in bright blue text to the terminal.
    """
    print(f"\033[96m{msg}\033[39m")

def query_user(msg: str, default: str='') -> str:
    """Prints a message in bright blue text to the terminal and returns user input,
    or the default value if the user provides no input.
    """
    resp = input(f"\033[96m{msg}\033[39m")
    
    return resp if resp != '' else default

def query_num(msg: str, default: int=None) -> int:
    ui.showCursor()
    try:
        resp = query_user(msg, str(default))
        resp = int(resp)
    except:
        if resp == '' and default is not None:
            return default
    
        ui.print_error(f"'{resp}' is not a valid number. Try again.")
        return query_num(msg, default)

    return resp
     
def query_date(msg: str, default: datetime=None) -> datetime:
    ui.showCursor()
    try:
        resp = query_user(msg)
        resp = utils.parse_datetime(resp)
    except BaseException:
        if resp == '' and default is not None:
            return default
        
        ui.print_error("The date entered is not valid. Try again.")
        return query_date(msg, default)
        
    return resp

def promptFile(filetypes) -> str:
    """
    Opens file explorer for the user to select a file and returns filepath.
    """
    return tkinter.filedialog.askopenfilename(filetypes=filetypes)

def promptDirectory() -> str:
    """
    Opens file explorer for the user to select a filepath for saving a file to."""
    return tkinter.filedialog.askdirectory()

def hideCursor():
    """
    Hide user cursor.
    """
    sys.stdout.write('\033[?25l')
    sys.stdout.flush()

def showCursor():
    """
    Unhide user cursor some time after calling hideCursor().
    """
    sys.stdout.write('\033[?25h')
    sys.stdout.flush()

def clear(version: str):
    """
    Clear terminal and display logotype."""
    os.system('cls' if os.name == 'nt' else 'clear')
    print(f""" _____        _        _____  _               _   
|  __ \\      | |      |  __ \\(_)             | |  
| |  | | __ _| |_ __ _| |  | |_ _ __ ___  ___| |_ 
| |  | |/ _` | __/ _` | |  | | | '__/ _ \\/ __| __|
| |__| | (_| | || (_| | |__| | | | |  __/ (__| |_ 
|_____/ \\__,_|\\__\\__,_|_____/|_|_|  \\___|\\___|\\__|

DataDirect copyright (c) AdDirect Incorporated 2025. Version {version}
==================================================================""")

def wait(s: int):
    """Wait for s seconds before executing rest of code.
    """
    time.sleep(s)

def pause(msg: str=None):
    """
    Wait for the user to press a key before executing rest of code.
    Print a string while paused with text."""
    if msg is not None:
        print(msg)

    return msvcrt.getch()

def exit():
    """Ends the program.
    """
    sys.exit()