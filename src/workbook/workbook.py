
class Workbook:
    path = "xl/workbook.xml"

    def __init__(self):
        self._sheets = []

    def add_worksheet(self, sheet):
        self._sheets.append(sheet)
