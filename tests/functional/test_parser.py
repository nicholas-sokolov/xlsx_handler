from unittest import TestCase

from src import load_workbook


class TestParser(TestCase):

    def setUp(self):
        self.workbook = load_workbook(r'..\xls_example\xlsm_file.xlsm')
        self.fixture_dict = {
            'x': 5,
            'y': 10,
            'z': 50,
            'total': 100,
            'vba': 100,
        }

    def test_get_value(self):
        worksheet = self.workbook['Sheet1']
        for row in range(1, worksheet.max_row):
            name = worksheet.cell(row, 1).value
            if name:
                value = worksheet.cell(row, 2).value
                self.assertEqual(value, self.fixture_dict[name])
