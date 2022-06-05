from tkinter.filedialog import *
import os

class CampointsExcelFile:
    """Retrieve excel file and store csv filename in the same path"""
    def __init__(self):
        self.filename_with_path = askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        self.___filename_without_extension = os.path.splitext(self.filename_with_path)[0]
        self.csv_export_filename = self.___filename_without_extension + ".csv"
        self.___filename = os.path.basename(self.filename_with_path)
        self.___filepath = os.path.dirname(self.filename_with_path)
