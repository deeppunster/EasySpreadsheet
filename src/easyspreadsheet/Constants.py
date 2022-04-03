"""
Constants.py - Constants, named tuples, and miscellaneous data.
"""
"""
Constants
"""

from typing import NamedTuple

SPREADSHEET_NAME = 'SimpleSpreadsheet.xlsx'
SPREADSHEET_NAME_2 = 'SimpleSpreadsheet2.xlsx'
SPREADSHEET_NAME_3 = 'SimpleSpreadsheet3.xlsx'

"""
sample test data for chart
"""

test_data = [
    ['Number', 'Batch 1', 'Batch 2'],
    [2, 30, 40],
    [3, 25, 40],
    [4, 30, 50],
    [5, 10, 30],
    [6, 5, 25],
    [7, 10, 50],
        [8, 50, 80]
]
nbr_tests = len(test_data)

"""
Named Tuples to manage parameterization
"""


class RefData(NamedTuple):
    min_col: int
    max_col: int
    min_row: int
    max_row: int
    titles_included: bool


class Area_3D_Chart_Info(NamedTuple):
    title: str
    style: int
    x_axis_title: str
    y_axis_title: str
    legend: str | None
    category_info: RefData
    data_info: RefData


# chart = AreaChart3D()
# chart.title = "Area Chart"
# chart.style = 13
# chart.x_axis.title = 'Test'
# chart.y_axis.title = 'Percentage'
# chart.legend = None

# cats = Reference(ws, min_col=1, min_row=1, max_row=7)
# data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7)
# chart.add_data(data, titles_from_data=True)
# chart.set_categories(cats)
