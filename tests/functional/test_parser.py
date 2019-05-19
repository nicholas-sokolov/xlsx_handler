import os
import shutil
from unittest import TestCase

from src import load_workbook

EXAMPLE_FILE_PATH = os.path.join(os.path.pardir, 'xls_example')
XLSM_ORIG_FILE_PATH = os.path.join(EXAMPLE_FILE_PATH, 'original.xlsm')
XLSM_TEMP_FILE_PATH = os.path.join(EXAMPLE_FILE_PATH, 'original_temp.xlsm')
XLSM_OUTPUT_FILE_PATH = os.path.join(EXAMPLE_FILE_PATH, 'output.xlsm')

xlsm_data_fixture = [
    ('x', 5),
    ('y', 10),
    ('z', 50),
    (),
    ('total', 100),
    ('vba', 100),
]


class TestParser(TestCase):

    def setUp(self):
        self.workbook = load_workbook(XLSM_ORIG_FILE_PATH)

    def tearDown(self):
        self.workbook.close()

    def test_get_value(self):
        worksheet = self.workbook['Sheet1']

        for row in range(1, worksheet.max_row):
            name = worksheet.cell(row, 1).value
            value = worksheet.cell(row, 2).value
            if xlsm_data_fixture[row-1]:
                expected_name, expected_value = xlsm_data_fixture[row-1]
                self.assertEqual(expected_name, name)
                self.assertEqual(expected_value, value)

    def test_set_value(self):
        if os.path.exists(XLSM_TEMP_FILE_PATH):
            os.remove(XLSM_TEMP_FILE_PATH)
        shutil.copy(XLSM_ORIG_FILE_PATH, XLSM_TEMP_FILE_PATH)

        expected_name = 'my app'
        expected_value = 100

        workbook = load_workbook(XLSM_TEMP_FILE_PATH)
        worksheet = workbook['Sheet1']
        worksheet.cell(7, 1).value = expected_name
        worksheet.cell(7, 2).value = expected_value
        workbook.save(XLSM_OUTPUT_FILE_PATH)

        workbook1 = load_workbook(XLSM_OUTPUT_FILE_PATH)
        worksheet1 = workbook1['Sheet1']

        name = worksheet1.cell(7, 1).value
        value = worksheet1.cell(7, 2).value

        self.assertEqual(name, expected_name)
        self.assertEqual(value, expected_value)