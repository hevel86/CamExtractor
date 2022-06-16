# Import necessary tools
from campoints_excel_file import CampointsExcelFile, DEGREES_COLUMN, POSITION_COLUMN
import pandas as pd
import openpyxl
from prettytable import PrettyTable
import matplotlib.pyplot as plt
import xlsxwriter

# Declare global constants
CSV_HEADER = '% MA_PERIODE=1 SL_PERIODE=1 CYCLIC=1'
CAMPOINTS_COLUMN = 'CAMPOINTS'
REGULAR_NUM_CAMPOINTS = 360
REGULAR_FINAL_CAMPOINT = 0
EXPECTED_ROTODEX_LIST_LENGTH = 37
ROTODEX_NUM_CAMPOINTS = 180
ROTODEX_FINAL_CAMPOINT = 1


def create_ascii_table(input_dictionary, table):
    """Create an ascii table from the derived data"""
    for list in input_dictionary:
        table.add_column(list.title(), input_dictionary[list])
    table.align = "l"


def create_graph(x_list, y_list):
    """Create graph using pyplot"""
    plt.plot(x_list, y_list)
    plt.title(cef.base_filename)
    plt.xlabel(POSITION_COLUMN.title())
    plt.ylabel(DEGREES_COLUMN.title())
    plt.get_current_fig_manager().set_window_title(cef.figure_title + " Motion Profile")
    plt.show()


def remove_null_from_dataframe(pandas_list):
    """Remove null values from list extracted via pandas library"""
    # Loop through the original and make changes to the new list
    clean_list = []
    for items in pandas_list:
        if pd.notna(items):
            clean_list.append(items)
    return clean_list


def panda_manipulation(excel_filename):
    """Use the pandas library to retrieve the degrees column and create a list"""
    # Extract data from the position and degrees column after removing the null values
    degrees_list = remove_null_from_dataframe(cef.dataframe[DEGREES_COLUMN].tolist())
    position_list = remove_null_from_dataframe(cef.dataframe[POSITION_COLUMN].tolist())

    # Check if it's a rotodex cam
    if len(position_list) <= EXPECTED_ROTODEX_LIST_LENGTH:
        cef.is_rotodex_cam = True

    # Check if the list ends at 355 or 360 for regular cams, or 175 or 180 for rotodex cams
    if not cef.is_rotodex_cam:
        num_campoints = REGULAR_NUM_CAMPOINTS
        final_degrees_position = REGULAR_FINAL_CAMPOINT
    else:
        num_campoints = ROTODEX_NUM_CAMPOINTS
        final_degrees_position = ROTODEX_NUM_CAMPOINTS

    if position_list[-1] != num_campoints:
        position_list.append(num_campoints)
        degrees_list.append(final_degrees_position)

    # Declare empty campoints list
    campoints_list = []
    # Divide by num campoints
    for campoints_index in degrees_list:
        # Don't add null items
        if pd.notna(campoints_index):
            campoints_list.append(campoints_index / num_campoints)

    # Create dictionary
    cef_dictionary = {
        POSITION_COLUMN: position_list,
        DEGREES_COLUMN: degrees_list,
        CAMPOINTS_COLUMN: campoints_list
    }
    return cef_dictionary


def export_xlsx():
    """Exports an XLSX version of the original file"""
    # Specify the output filename and path
    xlsx_filename = cef.filename_without_extension + ".xlsx"
    # Create the XLSX writer
    xlsx_writer = pd.ExcelWriter(xlsx_filename, engine="xlsxwriter")
    df = pd.DataFrame(cef_dict)
    df.to_excel(xlsx_writer, sheet_name="Values", index=False)
    xlsx_writer.save()


def export_csv(csv_campoints_list):
    """Use the pandas library to export the campoints to CSV"""
    # Pandas dataframe to export CSV
    df = pd.DataFrame(csv_campoints_list, columns=[CSV_HEADER])
    df.to_csv(cef.csv_export_filename, index=False)


# Create new instance of campoints Excel file
cef = CampointsExcelFile()

# Check for valid data in the Excel file
if cef.is_valid_data():
    # Get campoints from pandas
    cef_dict = panda_manipulation(cef.filename_with_path)

    # Make sure to use the list copy functionality if we need to edit the list
    # simply setting "equals to" makes it act as a pointer
    br_campoints = cef_dict[CAMPOINTS_COLUMN].copy()

    # Create a campoints table to print to the console
    cef_table = PrettyTable()
    create_ascii_table(cef_dict, cef_table)

    # Export CSV using pandas to_csv functionality
    export_csv(br_campoints)

    # Print the ascii table
    print(cef_table)
    create_graph(x_list=cef_dict[POSITION_COLUMN], y_list=cef_dict[DEGREES_COLUMN])