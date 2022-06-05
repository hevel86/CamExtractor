# Import necessary tools
from campoints_excel_file import CampointsExcelFile
import pandas as pd
from prettytable import PrettyTable

# Declare variables
CSV_HEADER = '% MA_PERIODE=1 SL_PERIODE=1 CYCLIC=1'
DEGREES_COLUMN = 'DEGREES'


def panda_manipulation(excel_filename):
    """Use the pandas library to retrieve the degrees column and create a list"""
    # Read Excel file in pandas
    df = pd.read_excel(excel_filename)
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


def export_csv(csv_campoints_list):
    """Use the pandas library to export the campoints to CSV"""
    print("Export CSV")
    # Pandas dataframe to export CSV
    df = pd.DataFrame(csv_campoints_list, columns=[CSV_HEADER])
    df.to_csv(cef.filename_without_extension + ".csv", index=False)


# Create new instance of campoints excel file
cef = CampointsExcelFile()

# Call functions
# Get full filename + path

# Get campoints from pandas + some editing
br_campoints = panda_manipulation(cef.filename_with_path)
# Export CSV using pandas to_csv functionality
export_csv(br_campoints)

# create a campoints table to print to the console
table = PrettyTable()
table.add_column("Cam Points", br_campoints)
table.align = "l"
print(table)









