def _get_column_letter(col_idx):
    """Convert a column number into a column letter (3 -> 'C')

    Right shift the column col_idx by 26 to find column letters in reverse
    order.  These numbers are 1-based, and can be converted to ASCII
    ordinals by adding 64.

    """
    # these indicies corrospond to A -> ZZZ and include all allowed
    # columns
    if not 1 <= col_idx <= 18278:
        raise ValueError("Invalid column index {0}".format(col_idx))
    letters = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx, 26)
        # check for exact division and borrow if needed
        if remainder == 0:
            remainder = 26
            col_idx -= 1
        letters.append(chr(remainder+64))
    return ''.join(reversed(letters))


def get_column_number(column_letter):
    col_idx = 1
    letters = []
    while col_idx < 26:
        if chr(col_idx+64) == column_letter:
            return col_idx
        col_idx += 1
    # TODO: this is just for first time. Workaround.
    raise Exception('invalid column')


class Cell:

    def __init__(self, row, column):
        self.row = row
        self.column = column
        self.type = None
        self._value = None
        self.formulae = None

    @property
    def column_latter(self):
        return _get_column_letter(self.column)

    @property
    def coordinate(self):
        return f'{self.column_latter}{self.row}'

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, item):
        try:
            self._value = int(item)
        except ValueError:
            try:
                self._value = float(item)
            except ValueError:
                self._value = item
                self.type = "inlineStr"
