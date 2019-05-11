from ..cell import Cell


class Worksheet:
    def __init__(self, name, path):
        self.name = name
        self.path = path
        self._rows = []

    @property
    def max_row(self):
        return len(self._rows)

    def add_row(self, row):
        self._rows.append(row)

    def get_row(self, row_number):
        for row in self._rows:
            if row.index == str(row_number):
                return row

    def cell(self, row, column):
        row_index = row - 1
        column_index = column - 1
        row = self.get_row(row_index)
        if row is None:
            return Cell()
        return row[column_index]
