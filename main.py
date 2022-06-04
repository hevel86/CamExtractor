# Import necessary tools
import pandas as pd
import xlrd
import os
import sys
from tkinter.filedialog import *
from prettytable import PrettyTable

# Declare variables
CSV_HEADER = '% MA_PERIODE=1 SL_PERIODE=1 CYCLIC=1'
DEGREES_COLUMN = 'DEGREES'
EXPORT_FILENAME = "test.csv"


def open_file():
    """Choose an excel file to extract the campoints from"""
    print("Open file")
    file_name = askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
    print(f"Filename: {filename}")
    return file_name


def panda_manipulation(excel_file):
    """Use the pandas library to retrieve the degrees column and create a list"""
    # Read Excel file in pandas
    df = pd.read_excel(filename)
    # Extract raw cam points from the DEGREES column
    raw_campoints = df[DEGREES_COLUMN].tolist()

    # Declare empty list
    campoints_list = []
    # Divide by 360 for regular cam
    for campoints_index in raw_campoints:
        campoints_list.append(campoints_index / 360.0)

    # Replace the last value with 0
    campoints_list[len(campoints_list) - 1] = 0
    # print(campoints_list)
    return campoints_list


def export_csv(csv_campoints):
    """Use the pandas library to export the campoints to CSV"""
    print("Export CSV")
    # Pandas dataframe to export CSV
    df = pd.DataFrame(csv_campoints, columns=[CSV_HEADER])
    df.to_csv(EXPORT_FILENAME, index=False)


# Call functions
# Get full filename + path
filename = open_file()
# Get campoints from pandas + some editing
br_campoints = panda_manipulation(filename)
# Export CSV using pandas to_csv functionality
export_csv(br_campoints)

# create a campoints table to print to the console
table = PrettyTable()
table.add_column("Cam Points", br_campoints)
table.align = "l"
print(table)









