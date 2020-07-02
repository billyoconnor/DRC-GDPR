#Script for handling data cleaning aspect of program
import os
import tkinter.messagebox
from tkinter import Tk
from tkinter.filedialog import askdirectory, askopenfilename
import numpy as np
import openpyxl
import pandas as pd
from pandas import read_excel

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
tkinter.messagebox.showinfo(title="Upload", message="Please choose the workbook you wish to upload")# Popup window explaining about to ask for source workbook
path1 = askopenfilename() # show an "Open" dialog box and return the path to the selected file - Store path as variable 'path1'


wb1 = pd.read_excel(path1)
wb1.loc[wb1['Template'].str.contains("Privacy"),'Type'] = 'Privacy'
wb1.loc[wb1['Template'].str.contains("Security"),'Type'] = 'Security'
wb1 = wb1.dropna(subset=['Type']) #remove the NaN values
wb1.reset_index(drop=True) # reset the index count which would have been mucked up

#Correct the duplicate names
tkinter.messagebox.showinfo(title="Upload", message="Please Select the Duplicate Mapping Workbook")# Popup window explaining about to ask for source workbook
path2 = askopenfilename() # show an "Open" dialog box and return the path to the selected file - Store path as variable 'path2'
wb2 = pd.read_excel(path2)
wb3 = wb1.merge(wb2[['ID', 'Concat']],on='ID', how='left') #Merging two tables on shared ID column (i.e performing a vlookup)
wb3['Concat'].fillna(wb3['Name'], inplace=True)
wb3['Concat2'] = wb3.Concat.map(str) + wb3.Type #works

tkinter.messagebox.showinfo(title="Upload", message="Please Select the Risk Export Workbook")# Popup window explaining about to ask for source workbook
path3 = askopenfilename() # show an "Open" dialog box and return the path to the selected file - Store path as variable 'path2'
wb4 = pd.read_excel(path3)

tkinter.messagebox.showinfo(title="Upload", message="Please Select the Risk to Question Mapping Workbook")# Popup window explaining about to ask for source workbook
path4 = askopenfilename() # show an "Open" dialog box and return the path to the selected file - Store path as variable 'path2'
wb5 = pd.read_excel(path4)
wb6 = wb4.merge(wb5[['Description', 'AssesmentType']],on='Description', how='left') #Merging two tables on shared ID column (i.e performing a vlookup)



