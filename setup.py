import sys
from cx_Freeze import setup, Executable

setup(
        name="spreadsheet ExcelUploader",
        version="1.0",
        description = "Python program to upload the result of filtering and parsing local xls file to google spreadsheet.",
        author = "beomjun-dev",
        executables = [Executable("spreadsheetExcelUploader.py")])