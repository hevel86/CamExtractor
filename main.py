# Import necessary tools
from campoints_excel_file import CampointsExcelFile, DEGREES_COLUMN, POSITION_COLUMN
import pandas as pd
import openpyxl
from prettytable import PrettyTable

# Declare global constants
CSV_HEADER = '% MA_PERIODE=1 SL_PERIODE=1 CYCLIC=1'
CAMPOINTS_COLUMN = 'CAMPOINTS'


def create_ascii_table(input_dictionary, table):
    """Create an ascii table from the derived data"""
    # Check if the cam points list has the same number of points
    while (len(input_dictionary[CAMPOINTS_COLUMN.lower()])) != len(input_dictionary[DEGREES_COLUMN.lower()]):
        input_dictionary[CAMPOINTS_COLUMN.lower()].pop(-1)
    for list in input_dictionary:
        table.add_column(list.title(), input_dictionary[list])
    table.align = "l"


def remove_null_from_dataframe(pandas_list):
    """Remove null values from list extracted via pandas library"""
    for index, item in pandas_list:
        if pd.isna(item):
            pandas_list[index] = 0
    return pandas_list


def panda_manipulation(excel_filename):
    """Use the pandas library to retrieve the degrees column and create a list"""
    # Extract data from the position and degrees column
    degrees_list = cef.dataframe[DEGREES_COLUMN].tolist()
    position_list = cef.dataframe[POSITION_COLUMN].tolist()

    # Declare empty list
    campoints_list = []
    # Divide by 360 for regular cam
    for campoints_index in degrees_list:
        # Don't add null items
        if pd.notna(campoints_index):
            campoints_list.append(campoints_index / 360.0)
    # Create dictionary
    cef_dictionary = {
        POSITION_COLUMN.lower(): position_list,
        DEGREES_COLUMN.lower(): degrees_list,
        CAMPOINTS_COLUMN.lower(): campoints_list
    }
    return cef_dictionary


def export_csv(csv_campoints_list):
    """Use the pandas library to export the campoints to CSV"""
    # Pandas dataframe to export CSV
    df = pd.DataFrame(csv_campoints_list, columns=[CSV_HEADER])
    df.to_csv(cef.csv_export_filename, index=False)


# Create new instance of campoints Excel file
cef = CampointsExcelFile()

# Check for valid data in the Excel file
if cef.is_valid_data:
    # Get campoints from pandas
    cef_dict = panda_manipulation(cef.filename_with_path)
    br_campoints = cef_dict[CAMPOINTS_COLUMN.lower()]

    # Check if the last position value is 360
    if cef_dict[POSITION_COLUMN.lower()][-1] != 360:
        br_campoints.append(0)

    # Export CSV using pandas to_csv functionality
    export_csv(br_campoints)

    # Create a campoints table to print to the console
    cef_table = PrettyTable()
    create_ascii_table(cef_dict, cef_table)
    print(cef_table)