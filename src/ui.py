from functools import partial
import os, sys, msvcrt, time
import tkinter
import tkinter.filedialog
from typing import Callable

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

def clear():
    """
    Clear the terminal."""
    os.system('cls' if os.name == 'nt' else 'clear')
    print("""
 _____        _        _____  _               _   
|  __ \\      | |      |  __ \\(_)             | |  
| |  | | __ _| |_ __ _| |  | |_ _ __ ___  ___| |_ 
| |  | |/ _` | __/ _` | |  | | | '__/ _ \\/ __| __|
| |__| | (_| | || (_| | |__| | | | |  __/ (__| |_ 
|_____/ \\__,_|\\__\\__,_|_____/|_|_|  \\___|\\___|\\__|

DataDirect copyright (c) AdDirect Incorporated 2025. Version 1.0.0
""")

def wait(s: int):
    """Wait for s seconds before executing rest of code.
    """
    time.sleep(s)

def pause():
    """
    Wait for the user to press a key before executing rest of code.
    Print a string while paused with text."""
    msvcrt.getch()

def on_error(cond: bool, message: str, terminate: bool=True):
    """Displays `message` in red on the interface if `cond` is True.
    Terminates the program if bool argument `terminate` is True.

    If `cond` is always True, then `raise_error` should be used instead.
    """
    if cond:
        clear()
        hideCursor()
        print(f'\033[31m{message} {'The program will now terminate. ' if terminate else ''}Press any key to continue.\033[39m')
        pause()
        showCursor()
        if terminate:
            exit()

raise_error = partial(on_error, cond=True)

def exit():
    """Ends the program.
    """
    sys.exit()
    showCursor()

class Menu():

    def __init__(self, title: str, *options: list['Option']):
        self._title = title
        self._options = options
    
    def open(self):
        # display options
        idx = 0
        num_options = len(self._options)
        print(self._title)
        print('\n' * num_options, end='')
        while True:
            sys.stdout.write(f'\033[{num_options}A')
            for i, option in enumerate(self._options):
                prefix = '-> ' if i == idx else '   '
                print(f'{prefix}{option.text}')

            key = msvcrt.getch()
            if key == b'H': # up arrow
                idx = (idx - 1) % len(self._options)
            elif key == b'P': # down arrow
                idx = (idx + 1) % len(self._options)
            elif key in [bytes(str(i), 'utf-8') for i in range(1, num_options + 1)]:
                # set index to pressed digit
                idx = int(key.decode()) - 1

                # redraw so that arrow is on the correct option
                sys.stdout.write(f'\033[{num_options}A')
                for i, option in enumerate(self._options):
                    prefix = '-> ' if i == idx else '   '
                    print(f'{prefix}{option.text}')

                # execute options command
                self._options[idx].command() if self._options[idx].command is not None else ''

                # return executed option
                return self._options[idx]
            
            elif key == b'\r': # enter
                # do the same thing as above, basically
                # unhide cursor
                sys.stdout.write('\033[?25h')
                sys.stdout.flush()

                # execute options command
                self._options[idx].command() if self._options[idx].command is not None else ''

                # return executed option
                return self._options[idx]

    @staticmethod
    def clear():
        os.system('cls' if os.name == 'nt' else 'clear')

class Option():
    def __init__(self, text: str, label: str, command: Callable=None):
        self.text = text
        self.label = label
        self.command = command