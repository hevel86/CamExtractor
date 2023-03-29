# Cam Extractor

This Python script is designed to extract data from a specified Excel file and perform operations on the data to generate CAMPOINTS data. The CAMPOINTS data is then exported to a CSV file and an XLSX file. The script also generates an ASCII table and a graph using pyplot.

## Prerequisites:
- Python 3.x
- pandas
- openpyxl
- prettytable
- matplotlib

## Usage:
1. Place the `campoints.py` file in a directory.
2. Ensure that the necessary libraries (pandas, openpyxl, prettytable, and matplotlib) are installed.
3. Place the Excel file that contains the data in the same directory.
4. Open the terminal and navigate to the directory where `campoints.py` and the Excel file are located.
5. Run the script using the following command: `python campoints.py`
6. The generated CSV and XLSX files will be stored in the same directory.

## Functionality:
- The `campoints_excel_file` module is imported to extract data from the Excel file.
- The `create_ascii_table` function is used to create an ASCII table from the derived data.
- The `create_graph` function is used to generate a graph using pyplot.
- The `remove_null_from_dataframe` function is used to remove null values from the extracted data.
- The `panda_manipulation` function uses pandas library to retrieve the degrees column and create a list of positions, degrees, and CAMPOINTS data.
- The `export_csv` function exports the CAMPOINTS data to a CSV file using the pandas library.
- The `export_xlsx` function exports an XLSX version of the original file.
- The `cef` object is created as an instance of the `CampointsExcelFile` class.
- The `is_valid_data` method is used to check for valid data in the Excel file.
- The `panda_manipulation` function is called to retrieve the CAMPOINTS data.
- The CAMPOINTS data is exported to a CSV file using the `export_csv` function.
- The ASCII table is generated using the `create_ascii_table` function.
- The graph is generated using the `create_graph` function.
