# Import other python tasks

# Import necessary tools
import pandas as pd
import xlrd
import openpyxl
import os
import sys
from tkinter.filedialog import *

# Declare variables
CSV_HEADER = '% MA_PERIODE=1 SL_PERIODE=1 CYCLIC=1'
DEGREES_COLUMN = 'DEGREES'
export_filename = "test.csv"

def openFile():
    print("Open file")
    filename = askopenfilename(filetypes=[("Excel files",".xlsx .xls")])
    print(filename)
    return filename

def pandaManipulation(excel_file):
    # Read excel file in pandas
    df = pd.read_excel(filename)
    # Extract raw campoints from the DEGREES column
    raw_campoints = df[DEGREES_COLUMN].tolist()
    #Declare empty list
    campoints_list = []
    # Divide by 360 for regular cam
    for campointsIndex in raw_campoints:
        campoints_list.append(campointsIndex / 360.0)

    # Replace the last value with 0
    campoints_list[len(campoints_list) - 1] = 0
    print(campoints_list)
    return campoints_list

def exportCSV(csv_campoints):
    print("Export CSV")
    # Pandas dataframe to export CSV
    df = pd.DataFrame(csv_campoints, columns=[CSV_HEADER])
    df.to_csv(export_filename, index=False)

# Call functions
try:
    # Get full filename + path
    filename = openFile()
    # Get campoints from pandas + some editing
    br_campoints = pandaManipulation(filename)
    # Export CSV using pandas to_csv functionality
    exportCSV(br_campoints)
except:
    print("Exception Error")









