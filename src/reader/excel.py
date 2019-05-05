from zipfile import ZipFile

from lxml import etree

from ..row import Row
from ..cell import Cell
from ..workbook.workbook import Workbook
from ..worksheet.worksheet import Worksheet


def find_file_path(file_list, pattern):
    file_path = None
    for file in file_list:
        if file.endswith(pattern):
            file_path = file
    return file_path


class ExcelReader:
    def __init__(self, filename):
        self._filename = filename
        self._zip_archive = ZipFile(self._filename, 'r')
        self._zip_files = self._zip_archive.namelist()
        self._shared_strings = []
        self._shared_values = {}
        self._rels = []
        self.workbook = Workbook()
        self.read()

    def read(self):
        self.read_relationship()
        self.read_shared_string()
        self.read_workbook()

    def read_relationship(self):
        relationship_path = find_file_path(self._zip_files, 'workbook.xml.rels')
        if not relationship_path:
            return
        relationship_raw = self._zip_archive.read(relationship_path)
        root = etree.fromstring(relationship_raw)
        nsmap = {k if k is not None else 'rel': v for k, v in root.nsmap.items()}
        for node in root.xpath('/rel:Relationships/rel:Relationship', namespaces=nsmap):
            rel = {}
            for key, value in node.attrib.items():
                rel[key] = value
            self._rels.append(rel)

    def read_shared_string(self):
        shared_strings_path = find_file_path(self._zip_files, 'sharedStrings.xml')
        shared_strings_raw = self._zip_archive.read(shared_strings_path)
        root = etree.fromstring(shared_strings_raw)
        nsmap = {k if k is not None else 'sh': v for k, v in root.nsmap.items()}
        for index, string in enumerate(root.xpath('/sh:sst/sh:si/sh:t', namespaces=nsmap)):
            self._shared_values[str(index)] = string.text

    def read_workbook(self):
        workbook_raw = self._zip_archive.read(self.workbook.path)
        root = etree.fromstring(workbook_raw)
        nsmap = {k if k is not None else 'wb': v for k, v in root.nsmap.items()}
        worksheets = {}
        for sheet in root.xpath('/wb:workbook/wb:sheets/wb:sheet', namespaces=nsmap):
            sheet_id = sheet.xpath('@r:id', namespaces=nsmap)[0]
            sheet_name = sheet.attrib['name']
            rel = self.find_relationship(sheet_id)
            sheet_filename = rel['Target']
            sheet_file_path = find_file_path(self._zip_files, sheet_filename)
            worksheets[sheet_name] = sheet_file_path
        self.read_worksheets(worksheets)

    def read_worksheets(self, worksheets):
        """

        :param dict worksheets:
        :return:
        """
        for ws_name, ws_path in worksheets.items():
            worksheet = Worksheet(ws_name, ws_path)

            worksheet_raw = self._zip_archive.read(ws_path)
            root = etree.fromstring(worksheet_raw)
            nsmap = {k if k is not None else 'ws': v for k, v in root.nsmap.items()}
            for row_node in root.xpath('./ws:sheetData/ws:row', namespaces=nsmap):

                row = Row(row_node.attrib['r'])
                for cell_node in row_node.xpath('ws:c', namespaces=nsmap):
                    cell = Cell()
                    cell.coordinate = cell_node.attrib['r']
                    cell.type = cell_node.attrib.get('t')

                    value_node = cell_node.find('ws:v', namespaces=nsmap)
                    if cell_node.attrib.get('t', False) == 's':
                        cell.value = self._shared_values.get(value_node.text)
                    else:
                        cell.value = value_node.text

                    if cell_node.find('ws:f', namespaces=nsmap) is not None:
                        cell.formulae = cell_node.find('ws:f', namespaces=nsmap).text

                    row.add_cell(cell)
                worksheet.add_row(row)

    def find_relationship(self, rel_id):
        for rel in self._rels:
            if rel['Id'] == rel_id:
                return rel
        return None


def load_workbook(filename):
    reader = ExcelReader(filename)
    return reader.workbook


