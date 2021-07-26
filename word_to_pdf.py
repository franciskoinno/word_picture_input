from docx2pdf import convert
import time
import os
from tkinter import Tk, filedialog


def select_folder():
    root = Tk() # pointing root to Tk() to use it as Tk() in program.
    root.withdraw() # Hides small tkinter window
    root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection
    open_file = filedialog.askdirectory() # Returns opened path as str
    return open_file


from_dir = select_folder()
from_arr = os.listdir(from_dir)

print(len(from_arr))

for folder in from_arr:
    convert(f"{from_dir}/{folder}/{folder}.docx", f"pdf/{folder}.pdf")
    time.sleep(1.2)

# comments