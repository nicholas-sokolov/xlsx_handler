class Cell:

    def __init__(self):
        self.coordinate = None
        self.type = None
        self._value = None
        self.formulae = None

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
