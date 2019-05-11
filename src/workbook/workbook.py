
class Workbook:
    path = "xl/workbook.xml"

    def __init__(self):
        self._sheets = []

    def __getitem__(self, item):
        for sheet in self._sheets:
            if sheet.name == item:
                return sheet

    def add_worksheet(self, sheet):
        self._sheets.append(sheet)

    def save(self):
        raise NotImplementedError
