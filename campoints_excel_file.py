from tkinter.filedialog import *
import os

class CampointsExcelFile:
    def __init__(self):
        self.filename_with_path = askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        self.filename_without_extension = os.path.splitext(self.filename_with_path)[0]
        self.filename = os.path.basename(self.filename_with_path)
        self.filepath = os.path.dirname(self.filename_with_path)

