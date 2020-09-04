'''
Created on Aug 15, 2020
@author: Alan-Hannan
Licence: MIT
'''

# ======================
# imports
# ======================
import tkinter as tk
from tkinter import Menu, filedialog
from tkinter import ttk
from tkinter import filedialog as fd
import pandas as pd
from fuzzywuzzy import process
import tkinter.messagebox



# ======================
# GlobalVariables
# ======================
idsFilePth = ""
lookupFilePth = ""


# ======================
# functions
# ======================

# load the path to the file with ids as a global variable
def _loadIDsFileButtonCommand():
    global idsFilePth
    name = fd.askopenfilename()
    idsFilePth = name
    print(idsFilePth)


# load the path to the file with lookup values as a global variable
def _loadLookupsFileButtonCommand():
    global lookupFilePth
    name = fd.askopenfilename()
    lookupFilePth = name
    print(lookupFilePth)


# Run the fuzzy lookup algorithm and export the output file
def _runFuzzyLookup():
    global idsFilePth
    global lookupFilePth
    dffinalresult = fuzzy(idsFilePth, lookupFilePth)
    dffinalresult.to_excel("output.xlsx")



# Exit GUI cleanly
def _quit():
    win.quit()  # win will exist when this function is called
    win.destroy()
    exit()

# Display Licence
def _about():
    tkinter.messagebox.showinfo('About:','Author: Alan Hannan \n Licence: MIT')


def fuzzy(idsFilePth, lookupFilePth):
    results = []
    dfids = pd.read_excel(idsFilePth)
    dflookup = pd.read_excel(lookupFilePth)
    twoDimentionalnpIds = dfids.to_numpy()
    lookupStringArray = twoDimentionalnpIds[:, 1]
    oneDimentionalLookuparrayOfArrays = dflookup.to_numpy()
    oneDimentionalLookuparray = oneDimentionalLookuparrayOfArrays[:, 0]
    for lookupValue in oneDimentionalLookuparray:
        result = []
        # print(lookupValue)
        result.append(lookupValue)
        matchTuple = process.extractOne(lookupValue, lookupStringArray)
        # print(matchTuple)
        result.append(str(matchTuple[0]))
        result.append(str(matchTuple[1]))
        # print(result)
        results.append(result)

    #print(results)
    dfresult = pd.DataFrame(results)
    dfresult.columns = ['Lookup', 'Name', 'MatchPercent']
    inner_join = pd.merge(dfresult, dfids , on='Name', how='inner')
    #print(inner_join)
    dffinalresult = inner_join[['Lookup', 'Name', 'id', 'MatchPercent']]
    #print(dffinalresult)
    return dffinalresult


# ======================
# procedural code
# ======================
# Create instance
win = tk.Tk()

# Add a title       
win.title("Office Minitools v0.1")
# ---------------------------------------------------------------
# Creating a Menu Bar
menuBar = Menu()
win.config(menu=menuBar)

# Add menu items
fileMenu = Menu(menuBar, tearoff=0)
fileMenu.add_command(label="New")
fileMenu.add_separator()
fileMenu.add_command(label="Exit", command=_quit)  # command callback
menuBar.add_cascade(label="File", menu=fileMenu)

# Add another Menu to the Menu Bar and an item
helpMenu = Menu(menuBar, tearoff=0)
helpMenu.add_command(label="About", command=_about)
menuBar.add_cascade(label="Help", menu=helpMenu)
# ---------------------------------------------------------------

# Tab Control / Notebook introduced here ------------------------
tabControl = ttk.Notebook(win)  # Create Tab Control

tab1 = ttk.Frame(tabControl)  # Create a tab
tabControl.add(tab1, text='FuzzyWuzzy')  # Add the tab

#tab2 = ttk.Frame(tabControl)  # Add a second tab
#tabControl.add(tab2, text='Levenshtein')  # Make second tab visible

tabControl.pack(expand=1, fill="both")  # Pack to make visible
# ---------------------------------------------------------------
# --------------------------------------------------------------------------FirstFrame
# We are creating a container frame to hold all other widgets
Load_Files_frame = ttk.LabelFrame(tab1, text='1-Load the Filled xlsx files ')

# using the tkinter grid layout manager
Load_Files_frame.grid(column=0, row=0, padx=8, pady=4)


# ---------------------------------------------------------------
# but frame won't be visible until we add widgets to it...
# Adding a Label
# ttk.Label(Load_Files_frame, text="Click:").grid(column=0, row=0, sticky='W')

# Adding a Button to load a file with Ids
ttk.Button(Load_Files_frame, text='Click to load file with ID values', command=_loadIDsFileButtonCommand).grid(column=0, row=1, sticky='W')
# Adding a Button to load a file with Lookups
ttk.Button(Load_Files_frame, text='Click to load file with lookup values', command=_loadLookupsFileButtonCommand).grid(column=0, row=2, sticky='W')

# -------------------------------------------------------------------------EndFirstFrame
# -------------------------------------------------------------------------SecondFrame
# We are creating a container frame to hold all other widgets
run_algorithm_frame = ttk.LabelFrame(tab1, text=' Run Fuzzy Lookup Algorithm ')
# using the tkinter grid layout manager
run_algorithm_frame.grid(column=0, row=1, padx=8, pady=4)

# ---------------------------------------------------------------
# but frame won't be visible until we add widgets to it...
# Adding a Label
#ttk.Label(run_algorithm_frame, text="Click:").grid(column=0, row=0, sticky='W')
ttk.Button(run_algorithm_frame, text='Click to Run the lookup algorithm', command=_runFuzzyLookup).grid(column=0, row=2, sticky='W')

# ------------------------------------------------------------------------EndSecondFrame

# ****************************************************************
# Cosmetics
# ****************************************************************
# Add some space around each label
for child in Load_Files_frame.winfo_children():
    #     child.grid_configure(padx=6, pady=6)
    #     child.grid_configure(padx=6, pady=3)
    child.grid_configure(padx=4, pady=2)  # adjust per visual taste of spacing around widgets

# ======================
# Start GUI
# ======================
win.mainloop()
