#!/usr/bin/python
# Python Libarys
import sys
import os
import string
import json
import io

# Libarys for working with Excel File
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

# Module files
from empty_excel import empty_excel_file

# Determine if Application is a Python Script or a complied .exe and define global _DIR
if getattr(sys, 'frozen', False):
    _DIR = os.path.dirname(sys.executable)
elif __file__:
    _DIR = os.path.dirname(__file__)

# Global Strings
XL_SHAREDSTRINGS_XML = "xl/sharedStrings.xml"
XL__RELS_WORKBOOK_XML_RELS = "xl/_rels/workbook.xml.rels"
_CONTENT_TYPES__XML = "[Content_Types].xml"


class excel_file():
    def __init__(self, name, dir=None, debug=False, empty=False):
        """
        name                                    | File Name without extension
        empty = False                           | Defines if a new Excel should be generated
        dir = "Path/were/Excel/is/located/"     | Defines the Path were the Excel is located
        debug = False                           | If true prints all Debug Texts
        """
        self._debug = debug
        self._name = name
        self._dir = dir or _DIR

        if empty:
            self.excel_contend, self.ZIP = self._create_new_empty_excel()
        else:
            self.excel_contend, self.ZIP = self._load_excel(
                f"{self._dir}/{self._name}.xlsx")

        # Open Workbook Rels
        self.workbook_rels = etree.fromstring(
            self.excel_contend[XL__RELS_WORKBOOK_XML_RELS])

        # Open Content Types
        self.content_types = etree.fromstring(
            self.excel_contend[_CONTENT_TYPES__XML])

        # Opens the all Sheets of the Excel
        self.sheets = []
        for n in [n for n in self.excel_contend.keys() if "xl/worksheets/sheet" in n]:
            tmp_sheet = etree.fromstring(self.excel_contend[n])
            tmp_data, tmp_data_index = [(child, i) for i, child in enumerate(
                tmp_sheet) if child.tag == "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData"][0]
            tmp_data = etree.fromstring(etree.tostring(tmp_data))
            self.sheets.append(
                {"sheet": tmp_sheet, "data": tmp_data, "index": tmp_data_index})

        # Shared Strings and List of it to check if the string is allready in the Shared Strings
        # No Default File, can be missing!
        if [n for n in self.excel_contend.keys() if XL_SHAREDSTRINGS_XML in n]:
            self.shared_strings = etree.fromstring(
                self.excel_contend[XL_SHAREDSTRINGS_XML])
            self.shared_strings_list = [v[0].text for v in self.shared_strings]

        else:
            self.shared_strings = None
            self.shared_strings_list = None

        # Opens all the Tables
        # No Default File, can be missing!
        self.tables = []
        tmp_tables = [n for n in self.excel_contend.keys()
                      if "xl/tables/table" in n]
        if tmp_tables:
            for n in tmp_tables:
                tmp_table = etree.fromstring(self.excel_contend[n])
                # ref could be "A1:J401"
                tmp_table_range = tmp_table.attrib["ref"]
                self.tables.append(
                    {"table": tmp_table, "range": tmp_table_range})

        if self._debug:
            print("* Loaded Excel")

    def save_as_folder(self):
        self._save_excel_memory()

        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, "a", ZIP_DEFLATED, False) as zip_file:
            for n, v in self.excel_contend.items():
                zip_file.writestr(n, v)

            # Exctract Excel to Folder
            zip_file.extractall(f"{self._dir}/{self._name}")

    def save_as_json(self):
        with open(f"{self._dir}/{self._name}.json", "w") as json_file:
            tmp_json = {n: v.decode() for n, v in self.excel_contend.items()}
            json.dump(tmp_json, json_file)

    def add_data(self, data, row, column, sheet):
        bounds_check = [x for i, x in zip(
            (row, column, sheet), ("Row", "Column", "Sheet")) if i < 1]
        if bounds_check:
            raise ValueError(f"{', '.join(bounds_check)} is too small")

        if type(data) in (type(12), type(12.)):
            # Can be added directly to the table
            data = str(float(data)).encode()

            # Add the Data to the Table
            self._add_data(data, row, column, sheet)

        else:
            # Check if the String is in the Shared Strings XML
            if self.shared_strings_list and data in self.shared_strings_list:
                member = True
                index = self.shared_strings_list.index(data)
            else:
                member = False

            # Needs to be added to shared Strings and then just indexed with a int starting from 0
            data = str(data).encode()

            if member:
                shared_strings_i = index
            else:
                shared_strings_i = self._add_to_shared_strings(data)

            data = str(shared_strings_i).encode()

            # Add the Data to the Table
            self._add_data(data, row, column, sheet, _str=True)

    def add_formula(self, data):
        """ Needs to be implemented. """
        pass

    def add_format(self):
        """ Needs to be implemented. """
        pass

    def add_sheet(self):
        """ Needs to be implemented. """
        pass

    def add_table(self):
        """ Needs to be implemented. """
        pass

    def add_image(self):
        """ Needs to be implemented. """
        pass

    def add_chart(self):
        """ Needs to be implemented. """
        pass

    def save_excel(self):
        """ Saves the file as .xlsx. """
        self._save_excel(self._name)

    def save_excel_at(self, path, name=None):
        """ Saves the file as .xlsx at given path. """
        self._save_excel(name or self._name, path=path)

    def _create_new_empty_excel(self):
        # The Empty Excel File Contend
        contend = empty_excel_file().get_encoded_dict()

        # Create ZipFile from contend
        with ZipFile(f"{self._dir}/{self._name}.xlsx", mode="w", compression=ZIP_DEFLATED) as new_ZIP:
            for items in contend.items():
                new_ZIP.writestr(*items)

        return contend, new_ZIP

    def _update_table_range(self, new_table_range):
        self.table_range = new_table_range

    def _add_to_shared_strings(self, data):
        """
        Neede becouse Strings dont get added to the Excel Sheet directly.
        They get stored ind a Shared Strings XML File and
        then they just get index like a list starting from (0: int).
        Its important to know that you need to add a self.sheet_1[4][row][column].attrib['t'] = "s"
        or it will not work (I handled this in self._add_data!)
        """

        # Check if the Shared Strings File exists
        if self.shared_strings is None:
            self.shared_strings = self._new_shared_strings()

        # Add ne Shared String Child
        child1 = etree.SubElement(self.shared_strings, "si")
        child2 = etree.SubElement(child1, "t")
        child2.text = data
        self.shared_strings_list = [v[0].text for v in self.shared_strings]

        return len(self.shared_strings) - 1

    def _new_shared_strings(self):
        empty_shared_strings = {
            XL_SHAREDSTRINGS_XML: "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\"></sst>"
        }
        # Extend Rels with Shared Strings
        tmp_rel_len = len([n for n in self.workbook_rels])
        tmp_rel = etree.SubElement(self.workbook_rels, "Relationship")
        tmp_rel.attrib["Id"] = f"rId{tmp_rel_len + 1}"
        tmp_rel.attrib["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        tmp_rel.attrib["Target"] = "sharedStrings.xml"

        # Add Shared Strings zo Content Types
        content_t = etree.SubElement(self.content_types, "Override")
        content_t.attrib["PartName"] = "/xl/sharedStrings.xml"
        content_t.attrib["ContentType"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"

        return etree.fromstring(empty_shared_strings[XL_SHAREDSTRINGS_XML].encode())

    def _get_column_name(self, index: int) -> str:
        if index < 1:
            raise ValueError("Index is too small")

        ascii_uppercase = string.ascii_uppercase
        ascii_uppercase_len = len(ascii_uppercase)

        if index <= ascii_uppercase_len:
            return ascii_uppercase[index - 1]

        result = ""
        idx = index
        while True:
            if idx > ascii_uppercase_len:
                idx, r = divmod(idx, ascii_uppercase_len)
                result = "".join([ascii_uppercase[r - 1], result])
            else:
                return "".join([ascii_uppercase[idx - 1], result])

    def _apply_shape(self, sheet_data: etree._Element, old_shape: list, new_shape: list):
        """ Iterate over the Data Sheet and add new Rows and Columns """
        # Check if there are rows in the Data
        if old_shape == [0, 0]:
            # Add first Row
            row_name, tmp_row = f"{1}", etree.SubElement(sheet_data, "row")
            tmp_row.attrib["r"] = row_name

            # Add first Column
            column = etree.SubElement(tmp_row, "c")
            column.attrib["r"] = f"{self._get_column_name(1)}{row_name}"

            old_shape = [1, 1]

        for i in range(new_shape[0]):
            # Add new row if outside old shape
            row_name = f"{i+1}"
            if i >= old_shape[0]:
                tmp_row = etree.SubElement(sheet_data, "row")
                tmp_row.attrib["r"] = row_name
            else:
                tmp_row = sheet_data[i]

            # Important iterate over every Row and Column!
            for ii in range(new_shape[1]):
                # Add new column if outside old shape
                if ii >= old_shape[1] or i >= old_shape[0]:
                    column = etree.SubElement(tmp_row, "c")
                    column.attrib["r"] = f"{self._get_column_name(ii+1)}{row_name}"

        return sheet_data, new_shape

    def _update_data_shape(self, sheet_data, old_shape, new_shape) -> tuple[etree._Element, list]:
        """ Updates the shape of the sheet_data """
        new_rows = 0 if old_shape[0] > new_shape[0] else new_shape[0] + \
            1 - old_shape[0]
        new_columns = 0 if old_shape[1] > new_shape[1] else new_shape[1] + \
            1 - old_shape[1]
        new_shape = [max([old_shape[0] + new_rows, new_shape[0]]),
                     max([old_shape[1] + new_columns, new_shape[1]])]

        if new_columns or new_rows:
            sheet_data, new_shape = self._apply_shape(
                sheet_data, old_shape, new_shape)

        return sheet_data, new_shape

    def _add_data(self, data, row, column, sheet, _str=False, table=None):
        """ Rows and Colums start with 1, 2, 3... """
        current_sheet = self.sheets[sheet - 1]
        tmp_sheet = current_sheet["sheet"]
        tmp_sheet_data = current_sheet["data"]
        tmp_sheet_data_index = current_sheet["index"]
        row_count = len(tmp_sheet_data.getchildren())
        if not row_count:
            data_shape = [0, 0]
        else:
            data_shape = [row_count, len(tmp_sheet_data[0].getchildren())]
        insert_shape = [row, column]

        # Update the shape of the Sheet Data if the new Data is outside the current Sheet
        tmp_sheet_data, new_shape = self._update_data_shape(
            tmp_sheet_data, data_shape, insert_shape)

        # Update Dimensions of Sheet and Active Cell
        # Dimensions of Sheet
        tmp_sheet[0].attrib["ref"] = f"{self._get_column_name(new_shape[1]-1)}{new_shape[0]}"

        # Add active cell
        if not len(list(tmp_sheet[1][0])):
            view = etree.SubElement(tmp_sheet[1][0], "selection")
            # Active Cell
            view.attrib["activeCell"] = f"{self._get_column_name(column)}{row+1}"
            # Active Cell
            view.attrib["sqref"] = f"{self._get_column_name(column)}{row+1}"
        else:
            # Active Cell
            tmp_sheet[1][0][0].attrib["activeCell"] = f"{self._get_column_name(column)}{row}"
            # Active Cell
            tmp_sheet[1][0][0].attrib["sqref"] = f"{self._get_column_name(column)}{row}"

        if self._debug:
            print(etree.tostring(tmp_sheet_data, pretty_print=True).decode())

        root = tmp_sheet_data[row - 1][column - 1]

        # Update new Table Size
        if table:
            self._update_table_range(
                f"A1:{self._get_column_name(new_shape[1])}{new_shape[0]}")

        # Check if there is a existing value, if so replace text and dont create new Sub Element
        if len(root):
            root[0].text = data
        else:
            value = etree.SubElement(root, "v")
            value.text = data

        # If it is a Shared String this Tag is needed!
        if _str:
            root.attrib['t'] = "s"

        # Update sheet from Data
        tmp_sheet.replace(tmp_sheet[tmp_sheet_data_index], tmp_sheet_data)

        # Save sheet
        self.sheets[sheet - 1] = {"sheet": tmp_sheet,
                                  "data": tmp_sheet_data, "index": tmp_sheet_data_index}

    def _load_excel(self, excel_file):
        input_zip = ZipFile(excel_file)
        return {name: input_zip.read(name) for name in input_zip.namelist()}, input_zip

    def _save_excel_memory(self):
        # Move Files back in Excel File
        self.excel_contend[XL__RELS_WORKBOOK_XML_RELS] = etree.tostring(self.workbook_rels,
                                                                        pretty_print=False, xml_declaration=True, encoding='UTF-8', standalone=True)

        """ Need to rebuild the Excel from all edited contend! """
        self.excel_contend[_CONTENT_TYPES__XML] = etree.tostring(self.content_types,
                                                                 pretty_print=False, xml_declaration=True, encoding='UTF-8', standalone=True)

        for i, sheet in enumerate(self.sheets):
            self.excel_contend[f"xl/worksheets/sheet{i + 1}.xml"] = etree.tostring(sheet["sheet"],
                                                                                   pretty_print=False, xml_declaration=True, encoding='UTF-8', standalone=True)

        if self.shared_strings is not None:
            self.excel_contend[XL_SHAREDSTRINGS_XML] = etree.tostring(self.shared_strings,
                                                                      pretty_print=False, xml_declaration=True, encoding='UTF-8', standalone=True)

        if self.tables:
            for i, table in enumerate(self.tables):
                self.excel_contend[f"xl/tables/table{i + 1}.xml"] = etree.tostring(table["table"],
                                                                                   pretty_print=False, xml_declaration=True, encoding='UTF-8', standalone=True)

    def _save_excel(self, name, _dir=None, path=None):
        # Get dir to save to
        _dir = _dir or self._dir
        self._save_excel_memory()

        if path:
            # , compression=ZIP_DEFLATED
            with ZipFile(f"{path}", mode="w", compression=ZIP_DEFLATED) as new:
                for n, v in self.excel_contend.items():
                    new.writestr(n, v)

        else:
            # , compression=ZIP_DEFLATED
            with ZipFile(f"{_dir}\\{name}.xlsx", mode="w", compression=ZIP_DEFLATED) as new:
                for n, v in self.excel_contend.items():
                    new.writestr(n, v)

        if self._debug:
            print("* Saved Excel")


if __name__ == "__main__":
    print("* Starting Test")
    _DIR = os.path.join(_DIR, "test_output")
    if not os.path.exists(_DIR):
        os.mkdir(_DIR)

    print("* Creating Empty Test File")
    test = excel_file("Test", empty=True)

    print("* Adding Data to Test File")
    test.add_data("This", 1, 1, 1)
    test.add_data("is", 2, 2, 1)
    test.add_data("a", 3, 3, 1)
    test.add_data("awesome", 4, 4, 1)
    test.add_data("Test", 5, 5, 1)

    print("* Save Test File as Excel")
    test.save_excel_at(f"{_DIR}/Test.xlsx")

    print("* Save Test File as Folder")
    test.save_as_folder()

    print("* Save Test File as Json")
    test.save_as_json()

    print("* Finished Test with no Errors")
