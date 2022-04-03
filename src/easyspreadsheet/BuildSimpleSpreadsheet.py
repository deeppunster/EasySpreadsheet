"""
BuildSimpleSpreadsheet.py - Build a simple spreadsheet from scratch.
"""
from pathlib import Path
from secrets import choice

from openpyxl import Workbook

#######################
# Constants
#######################
SPREADSHEET_NAME_2 = 'SimpleSpreadsheet2.xlsx'


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

    def rename_current_spreadsheet(self, new_name: str):
        """
        Rename the label for the current sheet.

        :param new_name:
        :return:
        """
        self.spreadsheet.title = new_name
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
    my_sheet.rename_current_spreadsheet('Python')
    my_sheet.add_cell('A1', 'Hi from Python!')

    # add some random numbers to a column
    column = 'c'
    my_sheet.add_cell(column + '1', 'Random Numbers')
    for row in range(2, 12):
        ref = column + str(row)
        contents = choice(range(1000))
        my_sheet.add_cell(ref, contents)

    my_sheet.save_spreadsheet(my_work_area, SPREADSHEET_NAME_2)
