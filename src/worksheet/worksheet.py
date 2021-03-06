from ..cell import Cell
from ..row import Row


class Worksheet:
    def __init__(self, workbook, name, path):
        self.parent = workbook
        self.name = name
        self.path = path
        self._rows = []
        self.__max_row = 0

    @property
    def max_row(self):
        return self.__max_row

    @max_row.setter
    def max_row(self, value):
        if self.max_row < value:
            self.__max_row = value

    def add_row(self, row):
        self._rows.append(row)
        self.max_row = int(row.index)

    def get_row(self, row_number):
        for row in self._rows:
            if row.index == str(row_number):
                return row

    def cell(self, row, column):
        column_index = column - 1
        row_item = self.get_row(row)
        if row_item is None:
            new_row = Row(str(row))
            new_cell = Cell(row, column)
            new_row.add_cell(new_cell)
            self.add_row(new_row)
            return new_cell
        try:
            return row_item[column_index]
        except IndexError:
            new_cell = Cell(row, column)
            row_item.add_cell(new_cell)
            return new_cell
