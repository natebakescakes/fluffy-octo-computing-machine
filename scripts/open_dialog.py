import tkinter
from tkinter import filedialog

def open_dialog():

    tkinter.Tk().withdraw() # Close the root window
    in_path = filedialog.askdirectory() # Choose folder
    # in_path = filedialog.askopenfilename # Choose filename (single)
    # in_path = filedialog.askopenfilename(multiple=True) # Choose filename (multiple)
    return in_path

if __name__ == "__main__":
    main()
