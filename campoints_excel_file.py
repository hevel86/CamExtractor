from tkinter.filedialog import *
import os
import pandas as pd

# Declare global constants
DEGREES_COLUMN = 'DEGREES'
POSITION_COLUMN = 'POSITION'
DEFAULT_TERMINAL_TEXT = '\033[0m'
RED_TERMINAL_TEXT = '\033[31m'
GREEN_TERMINAL_TEXT = '\033[32m'
BLUE_TERMINAL_TEXT = '\033[34m'


class CampointsExcelFile:
    """Retrieve Excel file and store csv filename in the same path"""

    def __init__(self):
        # Retrieve filename + UNC
        self.filename_with_path = askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        # Get the full UNC without the extension
        self.filename_without_extension = os.path.splitext(self.filename_with_path)[0]
        # Filename only, no extension, no path
        self.figure_title = os.path.basename(self.filename_with_path.split('.')[0])
        # Full UNC with csv extension
        self.csv_export_filename = self.filename_without_extension + ".csv"
        # Filename only with extension
        self.base_filename = os.path.basename(self.filename_with_path)
        self.___filepath = os.path.dirname(self.filename_with_path)
        self.dataframe = None
        self.is_rotodex_cam = False

    def is_valid_data(self):
        """Check if the uploaded Excel file is valid"""
        # Read Excel file in pandas

        # Assign the validity to false, initially
        file_validity = False
        # Check for a valid campoints sheet in the pandas dataframe
        try:
            xl = pd.ExcelFile(self.filename_with_path)
            # Loop through each sheet
            for sheet in xl.sheet_names:
                print(f'Scanning sheet {BLUE_TERMINAL_TEXT}{sheet}{DEFAULT_TERMINAL_TEXT}')
                xl.parse(sheet)
                df = pd.read_excel(xl, sheet)
                # Check if the data we're looking for exists
                if (DEGREES_COLUMN in df) and (POSITION_COLUMN in df):
                    self.dataframe = df

                    file_validity = True
                #else:
                #    print("Invalid excel file.  Exiting...")
                #    return False
            if file_validity:
                print(f"{GREEN_TERMINAL_TEXT}Valid{DEFAULT_TERMINAL_TEXT} excel file.  Continuing...")
            else:
                print(f"{RED_TERMINAL_TEXT}Invalid{DEFAULT_TERMINAL_TEXT} excel file. Exiting...")
            return file_validity
        except Exception as e:
            print(e)





