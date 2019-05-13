from zipfile import ZipFile, ZIP_DEFLATED
import shutil
import os
import tempfile

from lxml import etree


def remove_from_zip(zipfname, *filenames):
    tempdir = tempfile.mkdtemp()
    try:
        tempname = os.path.join(tempdir, 'new.zip')
        with ZipFile(zipfname, 'r') as zipread:
            with ZipFile(tempname, 'w') as zipwrite:
                for item in zipread.filelist:
                    if item.filename not in filenames:
                        data = zipread.read(item.filename)
                        zipwrite.writestr(item, data)
        shutil.move(tempname, zipfname)
    finally:
        shutil.rmtree(tempdir)


class Workbook:
    path = "xl/workbook.xml"

    def __init__(self, reader):
        self.reader = reader
        self._sheets = []

    def __getitem__(self, item):
        for sheet in self._sheets:
            if sheet.name == item:
                return sheet

    def add_worksheet(self, sheet):
        self._sheets.append(sheet)

    def save(self, filename):
        shutil.copy(self.reader.filename, filename)

        for worksheet in self._sheets:
            with ZipFile(filename, 'r', ZIP_DEFLATED, allowZip64=True) as f:
                with f.open(worksheet.path) as file:
                    data_raw = file.read()
            root = etree.fromstring(data_raw)
            nsmap = {k if k is not None else 'ws': v for k, v in root.nsmap.items()}

            row_root = root.find('./ws:sheetData', namespaces=nsmap)

            for row in worksheet._rows:
                if not root.xpath(f"//ws:row[@r='{row.index}']", namespaces=nsmap):
                    new_row = etree.SubElement(row_root, 'row', attrib={'r': row.index})
                    for cell in row._cells:
                        attrs = {'r': cell.coordinate}
                        if cell.type is not None:
                            attrs['t'] = cell.type
                        new_cell = etree.SubElement(new_row, 'c', attrib=attrs)
                        root_for_value = new_cell
                        if cell.type == 'inlineStr':
                            root_for_value = etree.SubElement(new_cell, 'is')
                            new_value = etree.SubElement(root_for_value, 't').text = str(cell.value).encode()
                        else:
                            new_value = etree.SubElement(root_for_value, 'v').text = str(cell.value).encode()
            print(etree.tostring(root, pretty_print=True).decode())
            remove_from_zip(filename, worksheet.path)

            with ZipFile(filename, 'a', ZIP_DEFLATED, allowZip64=True) as f:
                with f.open(worksheet.path, 'w') as file:
                    file.write(etree.tostring(root, pretty_print=True))

        raise NotImplementedError
