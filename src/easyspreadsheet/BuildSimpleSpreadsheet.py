"""
BuildSimpleSpreadsheet.py - Build a simple spreadsheet from scratch.
"""
from pathlib import Path

from openpyxl import Workbook

#######################
# Constants
#######################
SPREADSHEET_NAME = 'SimpleSpreadsheet.xlsx'


class ManageSpreadsheet:
    """
    Provide the functions for managing a spreadsheet.
    """

    def __init__(self):
        self.workbook = None
        self.spreadsheet = None
        return

    def create_spreadsheet(self):
        """
        Create a spreadsheet.

        :return:
        """
        self.workbook = Workbook()
        self.spreadsheet = self.workbook.active
        return

    def add_cell(self, cell_location: str, value):
        """
        Add a value at an absolute cell location.

        :param cell_location: absolute cell reference e.g. 'A1' or 'GG297'
        :param value: string, number, or formula
        :return:
        """
        self.spreadsheet[cell_location] = value
        return

    def save_spreadsheet(self, filepath: Path, filename: str):
        """
        Save the spreadsheet to a specified location.

        :param filepath: Fully qualified path to the directory
        :param filename: file name for the spreadsheet
        :return:
        """
        file_loc: Path = filepath
        file_name: str = filename
        full_path = file_loc / file_name
        self.workbook.save(full_path)
        return


if __name__ == '__main__':
    my_work_area = Path.cwd()
    my_sheet = ManageSpreadsheet()
    my_sheet.create_spreadsheet()
    my_sheet.add_cell('A1', 'Hi from Python!')
    my_sheet.save_spreadsheet(my_work_area, SPREADSHEET_NAME)
