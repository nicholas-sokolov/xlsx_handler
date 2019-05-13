from .cell import Cell
from .reader.excel import logging


class Row:
    def __init__(self, index):
        self.index = index
        self._cells = []

    def __getitem__(self, item):
        # try:
        return self._cells[item]
        # except IndexError:
        #     logging.debug('No such column')
        #     return Cell()

    def add_cell(self, cell):
        self._cells.append(cell)
