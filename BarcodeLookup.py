import tkinter as tk
from zebra import Zebra
from tkinter import ttk
import os
import pandas as pd
from tkinter import *

# The reason for this code:
# Products were checked in and were never physically given the unique Product ID.
# This file reads an exported inventory file and will return the unique Product ID based on the unique criteria that product was given.

username = os.path.join(os.path.expandvars('%userprofile%')) # Return Userprofile
File = 'example.xlsx' # Read this file in directory
columns = ['PrintValue', 'BarcodeScanned'] # Setting variable of columns we want to read
df = pd.read_excel(File, usecols=columns) # Calling pandas to create a dataframe from reading the file
app = Tk()
data = []

# This function is to exit the program.
def close():
    app.destroy()

# This function will search the barcode you've scanned in column and return the corresponding value
# This is where we also utilize the zebra library to create a template for the thermal printer to print out
def print_trg(self):
    if self in df['BarcodeScanned'].values:
        df2 = df.loc[df['BarcodeScanned'] == str(self), 'PrintValue']
        new = df2.to_list()
        for i in new:
            label1 = ("^XA~TA000~JSN^LT0^MNW^MTD^PON^PMN^LH0,0^JMA^PR5,5~SD15^JUS^LRN^CI0^XZ"
                      "^XA"
                      "^MMT"
                      "^PW609"
                      "^LL0406"
                      "^LS0"
                      "^BY3,3,154^FT594,234^BCI,,Y,N"
                      "^FD>:Value->512345678>69^FS"
                      "^PQ1,0,1,Y^XZ")
            new1 = label1.replace('^FD>:Value->512345678>69^FS', '^FD>:' + i + '^FS')
            z = Zebra('ZDesigner GK420d')
            z.output(new1)
            return i


# This function is setting argumaents on the barcode, if you want to add specific text in front of the barcode value.
def limit(*args):
    value = BarcodeScanned.get()
    if len(value) == 13: # This 13 value can be changed based on the length of your barcode numbers
        BarcodeScanned.set(value[0:0])
        lst = ['Barcode Scanned: ' + value, ' Printing Value: ' + str(print_trg(value))]
        s = data.append(lst)
        set_list(s)


# This is where we develop a listbox within the tkinter library, it appends values in the listbox.
# This helps us keep track of all the barcodes we have scanned recently.
def get_listbox(container, height=None, width=None):
    global w
    sb = tk.Scrollbar(container, orient=tk.VERTICAL)
    w = tk.Listbox(container, relief=tk.GROOVE, selectmode=MULTIPLE, height=height, width=width,
                   background='white',
                   font='TkFixedFont',
                   yscrollcommand=sb.set, )
    sb.config(command=w.yview)
    w.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
    sb.pack(fill=tk.Y, expand=1)
    return w

# This is how we write to the tkinter listbox by specifying the labels of the frame and what kind of input.
BarcodeScanned = StringVar()
BarcodeScanned.trace('w', limit)
f = ttk.Frame()
ttk.Label(f, text="(Case Sensitive)", font=("Arial", 10)).pack()
ttk.Label(f, text=" Scan label: ", font=("Arial", 20)).pack()
ttk.Entry(f, textvariable=BarcodeScanned, width=15, font='Arial 20').pack(padx=5, pady=5)
lstItems = get_listbox(f, 40, 80)


# It's always nice to have a reset button.
# This restarts the program if you've updated the file with new data.
def restart():
    import sys
    import os
    """Restarts the current program.
        Note: this function does not return. Any cleanup action (like
        saving data) must be done before calling this function."""
    python = sys.executable
    os.execl(python, python, *sys.argv)


def set_list(s):
    lstItems.delete(0, tk.END)
    for key, value in data:
        lstItems.insert(tk.END, tuple([key, value]))
        pass


# Finally, we start packing all the features of tkinter and the functions.
# Added the buttons to restart the program and to exit the program.
w = ttk.Frame()
ttk.Button(w, text="Restart Program", command=restart).pack(pady=70)
ttk.Button(w, text="Exit Program", command=close).pack()
f.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
w.pack(side=tk.RIGHT, fill=tk.BOTH, expand=1)

f.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
w.pack(side=tk.RIGHT, fill=tk.BOTH, expand=1)

app.mainloop()
