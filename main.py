# Import necessary tools
from campoints_excel_file import CampointsExcelFile
import pandas as pd
import openpyxl
from prettytable import PrettyTable

# Declare variables
CSV_HEADER = '% MA_PERIODE=1 SL_PERIODE=1 CYCLIC=1'
DEGREES_COLUMN = 'DEGREES'
POSITION_COLUMN = 'POSITION'
CAMPOINTS_COLUMN = 'CAMPOINTS'


def create_ascii_table(input_dictionary, table):
    """Create an ascii table from the derived data"""
    # Check if the cam points list has the same number of points
    while (len(input_dictionary[CAMPOINTS_COLUMN.lower()])) != len(input_dictionary[DEGREES_COLUMN.lower()]):
        input_dictionary[CAMPOINTS_COLUMN.lower()].pop(-1)
    table.add_column(POSITION_COLUMN.title(), input_dictionary[POSITION_COLUMN.lower()])
    table.add_column(DEGREES_COLUMN.title(), input_dictionary[DEGREES_COLUMN.lower()])
    table.add_column(CAMPOINTS_COLUMN.title(), input_dictionary[CAMPOINTS_COLUMN.lower()])
    table.align = "l"


def remove_null_from_dataframe(pandas_list):
    """Remove null values from list extracted via pandas library"""
    for index, item in pandas_list:
        if pd.isna(item):
            pandas_list[index] = 0
    return pandas_list


def panda_manipulation(excel_filename):
    """Use the pandas library to retrieve the degrees column and create a list"""
    # Read Excel file in pandas
    xl = pd.ExcelFile(excel_filename)
    # Loop through each sheet
    for sheet in xl.sheet_names:
        xl.parse(sheet)
        df = pd.read_excel(xl, sheet)
        # Check if the data we're looking for exists
        if DEGREES_COLUMN in df:
            # Extract data from the position and degrees column
            degrees_list = df[DEGREES_COLUMN].tolist()
            position_list = df[POSITION_COLUMN].tolist()

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
    # TODO Add check for valid data


def export_csv(csv_campoints_list):
    """Use the pandas library to export the campoints to CSV"""
    # Pandas dataframe to export CSV
    df = pd.DataFrame(csv_campoints_list, columns=[CSV_HEADER])
    df.to_csv(cef.csv_export_filename, index=False)


# Create new instance of campoints Excel file
cef = CampointsExcelFile()

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