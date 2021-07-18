#!/usr/bin/python
# Python Libarys
import sys, os, string, json, io

# Libarys for working with Excel File
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

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
            self.excel_contend, self.ZIP = self._load_excel(f"{self._dir}/{self._name}.xlsx")

        # Open Workbook Rels
        self.workbook_rels = etree.fromstring(self.excel_contend[XL__RELS_WORKBOOK_XML_RELS])

        # Open Content Types
        self.content_types = etree.fromstring(self.excel_contend[_CONTENT_TYPES__XML])

        # Opens the all Sheets of the Excel
        self.sheets = []
        for n in [n for n in self.excel_contend.keys() if "xl/worksheets/sheet" in n]:
            tmp_sheet = etree.fromstring(self.excel_contend[n])
            tmp_data, tmp_data_index = [(child, i) for i, child in enumerate(tmp_sheet) if child.tag == "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData"][0]
            tmp_data = etree.fromstring(etree.tostring(tmp_data))
            self.sheets.append({"sheet": tmp_sheet, "data": tmp_data, "index": tmp_data_index})

        # Shared Strings and List of it to check if the string is allready in the Shared Strings
        # No Default File, can be missing!
        if [n for n in self.excel_contend.keys() if XL_SHAREDSTRINGS_XML in n]:
            self.shared_strings =  etree.fromstring(self.excel_contend[XL_SHAREDSTRINGS_XML])
            self.shared_strings_list =  [v[0].text for v in self.shared_strings]

        else:
            self.shared_strings =  None
            self.shared_strings_list =  None

        # Opens all the Tables
        # No Default File, can be missing!
        self.tables = []
        tmp_tables = [n for n in self.excel_contend.keys() if "xl/tables/table" in n]
        if tmp_tables:
            for n in tmp_tables:
                tmp_table = etree.fromstring(self.excel_contend[n])
                tmp_table_range = tmp_table.attrib["ref"] # ref could be "A1:J401"
                self.tables.append({"table": tmp_table, "range": tmp_table_range})                

        if self._debug: print("* Loaded Excel")
    
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
        bounds_check = [x for i, x in zip((row, column, sheet), ("Row", "Column", "Sheet")) if i < 1]
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
        """
        File Structure of Empty Excel (.xlsx 2019)
        Excel_Name/
            [Content_Types].xml
            _rels/
                .rels
            docProps/
                app.xml
                core.xml
            xl/
                styles.xml
                workbook.xml
                _rels/
                    workbook.xml.rels
                theme/
                    theme1.xml
                worksheets/
                    sheet1.xml
        """
        # The Empty Excel File Contend, just opened .xlsx as .zip and saved in Dict as .json
        empty_excel_contend_2019 = {
            _CONTENT_TYPES__XML: "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>",
            "_rels/.rels": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>",
            "xl/workbook.xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x15 xr xr6 xr10 xr2\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\"><fileVersion appName=\"xl\" lastEdited=\"7\" lowestEdited=\"7\" rupBuild=\"22730\"/><workbookPr defaultThemeVersion=\"166925\"/><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x15\"><x15ac:absPath url=\"C:\\Users\\reute\\Documents\\Workspace\\Excel\\Excel_Hack\\1.01\\\" xmlns:x15ac=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac\"/></mc:Choice></mc:AlternateContent><xr:revisionPtr revIDLastSave=\"0\" documentId=\"8_{9FF946F3-D4C2-465D-8F28-F8B3A6C33DB3}\" xr6:coauthVersionLast=\"45\" xr6:coauthVersionMax=\"45\" xr10:uidLastSave=\"{00000000-0000-0000-0000-000000000000}\"/><bookViews><workbookView xWindow=\"-120\" yWindow=\"-120\" windowWidth=\"29040\" windowHeight=\"18240\" xr2:uid=\"{35C2526F-4B0F-4919-B27C-93DF7AE2D330}\"/></bookViews><sheets><sheet name=\"Tabelle1\" sheetId=\"1\" r:id=\"rId1\"/></sheets><calcPr calcId=\"191029\"/><extLst><ext uri=\"{140A7094-0E35-4892-8432-C4D2E57EDEB5}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:workbookPr chartTrackingRefBase=\"1\"/></ext><ext uri=\"{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}\" xmlns:xcalcf=\"http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures\"><xcalcf:calcFeatures><xcalcf:feature name=\"microsoft.com:RD\"/><xcalcf:feature name=\"microsoft.com:Single\"/><xcalcf:feature name=\"microsoft.com:FV\"/><xcalcf:feature name=\"microsoft.com:CNMTM\"/></xcalcf:calcFeatures></ext></extLst></workbook>",
            XL__RELS_WORKBOOK_XML_RELS: "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>",
            "xl/worksheets/sheet1.xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" xr:uid=\"{D84DA2E2-49BD-44CF-A604-CEA970C260AE}\"><dimension ref=\"A1\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews><sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/><sheetData/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.78740157499999996\" bottom=\"0.78740157499999996\" header=\"0.3\" footer=\"0.3\"/></worksheet>",
            "xl/theme/theme1.xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2><a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2><a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4><a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6><a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"\u6e38\u30b4\u30b7\u30c3\u30af Light\"/><a:font script=\"Hang\" typeface=\"\ub9d1\uc740 \uace0\ub515\"/><a:font script=\"Hans\" typeface=\"\u7b49\u7ebf Light\"/><a:font script=\"Hant\" typeface=\"\u65b0\u7d30\u660e\u9ad4\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/><a:font script=\"Armn\" typeface=\"Arial\"/><a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/><a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/><a:font script=\"Java\" typeface=\"Javanese Text\"/><a:font script=\"Lisu\" typeface=\"Segoe UI\"/><a:font script=\"Mymr\" typeface=\"Myanmar Text\"/><a:font script=\"Nkoo\" typeface=\"Ebrima\"/><a:font script=\"Olck\" typeface=\"Nirmala UI\"/><a:font script=\"Osma\" typeface=\"Ebrima\"/><a:font script=\"Phag\" typeface=\"Phagspa\"/><a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Sora\" typeface=\"Nirmala UI\"/><a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/><a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/><a:font script=\"Tfng\" typeface=\"Ebrima\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"\u6e38\u30b4\u30b7\u30c3\u30af\"/><a:font script=\"Hang\" typeface=\"\ub9d1\uc740 \uace0\ub515\"/><a:font script=\"Hans\" typeface=\"\u7b49\u7ebf\"/><a:font script=\"Hant\" typeface=\"\u65b0\u7d30\u660e\u9ad4\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/><a:font script=\"Armn\" typeface=\"Arial\"/><a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/><a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/><a:font script=\"Java\" typeface=\"Javanese Text\"/><a:font script=\"Lisu\" typeface=\"Segoe UI\"/><a:font script=\"Mymr\" typeface=\"Myanmar Text\"/><a:font script=\"Nkoo\" typeface=\"Ebrima\"/><a:font script=\"Olck\" typeface=\"Nirmala UI\"/><a:font script=\"Osma\" typeface=\"Ebrima\"/><a:font script=\"Phag\" typeface=\"Phagspa\"/><a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Sora\" typeface=\"Nirmala UI\"/><a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/><a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/><a:font script=\"Tfng\" typeface=\"Ebrima\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E510DB2}\"><thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\"/></a:ext></a:extLst></a:theme>",
            "xl/styles.xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac x16r2 xr\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\"><fonts count=\"1\" x14ac:knownFonts=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Standard\" xfId=\"0\" builtinId=\"0\"/></cellStyles><dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/><extLst><ext uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/></ext><ext uri=\"{9260A510-F301-46a8-8635-F512D64BE5F5}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\"/></ext></extLst></styleSheet>",
            "docProps/core.xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator>Philipp Reuter</dc:creator><cp:lastModifiedBy>Philipp Reuter</cp:lastModifiedBy><dcterms:created xsi:type=\"dcterms:W3CDTF\">2020-05-15T16:44:06Z</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">2020-05-15T16:44:21Z</dcterms:modified></cp:coreProperties>",
            "docProps/app.xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><TotalTime>0</TotalTime><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Arbeitsbl\u00e4tter</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=\"1\" baseType=\"lpstr\"><vt:lpstr>Tabelle1</vt:lpstr></vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>"
        }
        #Encode the Strings
        empty_excel_contend_2019 = {n: v.encode() for n, v in empty_excel_contend_2019.items()}

        with ZipFile(f"{self._dir}/{self._name}.xlsx", mode="w", compression=ZIP_DEFLATED) as new_ZIP:
            for n, v in empty_excel_contend_2019.items():
                new_ZIP.writestr(n, v)

        return empty_excel_contend_2019, new_ZIP

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
        if self.shared_strings is None: self.shared_strings = self._new_shared_strings()

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
        new_rows = 0 if old_shape[0] > new_shape[0] else new_shape[0] + 1 - old_shape[0]
        new_columns = 0 if old_shape[1] > new_shape[1] else new_shape[1] + 1 - old_shape[1]
        new_shape = [max([old_shape[0] + new_rows, new_shape[0]]), max([old_shape[1] + new_columns, new_shape[1]])]

        if new_columns or new_rows:
            sheet_data, new_shape = self._apply_shape(sheet_data, old_shape, new_shape)

        return sheet_data, new_shape

    def _add_data(self, data, row, column, sheet, _str=False, table=None):
        """ Rows and Colums start with 1, 2, 3... """
        current_sheet = self.sheets[sheet - 1]
        tmp_sheet = current_sheet["sheet"]
        tmp_sheet_data = current_sheet["data"]
        tmp_sheet_data_index = current_sheet["index"]
        row_count = len(tmp_sheet_data.getchildren())
        if not row_count: data_shape = [0, 0]
        else:
            data_shape = [row_count, len(tmp_sheet_data[0].getchildren())]
        insert_shape = [row, column]

        # Update the shape of the Sheet Data if the new Data is outside the current Sheet
        tmp_sheet_data, new_shape = self._update_data_shape(tmp_sheet_data, data_shape, insert_shape)

        # Update Dimensions of Sheet and Active Cell
        tmp_sheet[0].attrib["ref"] = f"{self._get_column_name(new_shape[1]-1)}{new_shape[0]}" # Dimensions of Sheet

        # Add active cell
        if not len(list(tmp_sheet[1][0])):
            view = etree.SubElement(tmp_sheet[1][0], "selection")
            view.attrib["activeCell"] = f"{self._get_column_name(column)}{row+1}" # Active Cell 
            view.attrib["sqref"] = f"{self._get_column_name(column)}{row+1}" # Active Cell
        else:
            tmp_sheet[1][0][0].attrib["activeCell"] = f"{self._get_column_name(column)}{row}" # Active Cell 
            tmp_sheet[1][0][0].attrib["sqref"] = f"{self._get_column_name(column)}{row}" # Active Cell

        if self._debug: print(etree.tostring(tmp_sheet_data, pretty_print=True).decode())

        root = tmp_sheet_data[row - 1][column - 1]

        # Update new Table Size
        if table: self._update_table_range(f"A1:{self._get_column_name(new_shape[1])}{new_shape[0]}")

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
        self.sheets[sheet - 1] = {"sheet": tmp_sheet, "data": tmp_sheet_data, "index": tmp_sheet_data_index}

    def _load_excel(self, excel_file):
        input_zip=ZipFile(excel_file)
        return {name: input_zip.read(name) for name in input_zip.namelist()}, input_zip

    def _save_excel_memory(self):
        # Move Files back in Excel File
        self.excel_contend[XL__RELS_WORKBOOK_XML_RELS] = etree.tostring(self.workbook_rels,
        pretty_print = False, xml_declaration = True, encoding='UTF-8', standalone=True)

        """ Need to rebuild the Excel from all edited contend! """
        self.excel_contend[_CONTENT_TYPES__XML] = etree.tostring(self.content_types,
            pretty_print = False, xml_declaration = True, encoding='UTF-8', standalone=True)

        for i, sheet in enumerate(self.sheets):
            self.excel_contend[f"xl/worksheets/sheet{i + 1}.xml"] = etree.tostring(sheet["sheet"],
            pretty_print = False, xml_declaration = True, encoding='UTF-8', standalone=True)
        
        if self.shared_strings is not None:
            self.excel_contend[XL_SHAREDSTRINGS_XML] = etree.tostring(self.shared_strings,
            pretty_print = False, xml_declaration = True, encoding='UTF-8', standalone=True)
        
        if self.tables:
            for i, table in enumerate(self.tables):
                self.excel_contend[f"xl/tables/table{i + 1}.xml"] = etree.tostring(table["table"],
                pretty_print = False, xml_declaration = True, encoding='UTF-8', standalone=True)


    def _save_excel(self, name, _dir=None, path=None):
        # Get dir to save to
        _dir = _dir or self._dir

        self._save_excel_memory()

        if path:
            with ZipFile(f"{path}", mode="w", compression=ZIP_DEFLATED) as new: #, compression=ZIP_DEFLATED
                for n, v in self.excel_contend.items():
                    new.writestr(n, v)

        else:
            with ZipFile(f"{_dir}\\{name}.xlsx", mode="w", compression=ZIP_DEFLATED) as new: #, compression=ZIP_DEFLATED
                for n, v in self.excel_contend.items():
                    new.writestr(n, v)

        if self._debug: print("* Saved Excel")


if __name__ == "__main__":
    print("* Starting Test")
    _DIR = os.path.join(_DIR, "test_output")
    if not os.path.exists(_DIR): os.mkdir(_DIR)

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
