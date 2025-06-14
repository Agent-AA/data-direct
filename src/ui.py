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
    Clear terminal and display logotype."""
    os.system('cls' if os.name == 'nt' else 'clear')
    print(""" _____        _        _____  _               _   
|  __ \\      | |      |  __ \\(_)             | |  
| |  | | __ _| |_ __ _| |  | |_ _ __ ___  ___| |_ 
| |  | |/ _` | __/ _` | |  | | | '__/ _ \\/ __| __|
| |__| | (_| | || (_| | |__| | | | |  __/ (__| |_ 
|_____/ \\__,_|\\__\\__,_|_____/|_|_|  \\___|\\___|\\__|

DataDirect copyright (c) AdDirect Incorporated 2025. Version 1.0.0
==================================================================""")

def wait(s: int):
    """Wait for s seconds before executing rest of code.
    """
    time.sleep(s)

def pause():
    """
    Wait for the user to press a key before executing rest of code.
    Print a string while paused with text."""
    msvcrt.getch()

def exit():
    """Ends the program.
    """
    sys.exit()