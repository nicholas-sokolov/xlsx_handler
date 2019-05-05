

class Worksheet:
    def __init__(self, name, path):
        self.name = name
        self.path = path
        self._rows = []

    def add_row(self, row):
        self._rows.append(row)
