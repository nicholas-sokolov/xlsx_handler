

class Row:
    def __init__(self, index):
        self.index = index
        self._cells = []

    def add_cell(self, cell):
        self._cells.append(cell)