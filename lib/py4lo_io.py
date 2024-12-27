#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2024 J. Férard <https://github.com/jferard>
#
#     This file is part of Py4LO.
#
#     Py4LO is free software: you can redistribute it and/or modify
#     it under the terms of the GNU General Public License as published by
#     the Free Software Foundation, either version 3 of the License, or
#     (at your option) any later version.
#
#     Py4LO is distributed in the hope that it will be useful,
#     but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#     GNU General Public License for more details.
#
#     You should have received a copy of the GNU General Public License
#     along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""
A module to handle I/O:
* read and write data to LO spreadsheets
* create import/export filters for CSV files
"""
# mypy: disable-error-code="import-untyped,import-not-found"
import csv
import encodings
import locale
import sys
from datetime import (date, datetime, time)
from enum import IntEnum, Enum
from typing import (Any, Callable, List, Iterator, Optional, Mapping, Tuple,
                    Iterable, cast)

# values of cell_typing
from py4lo_commons import uno_path_to_url
from py4lo_helper import (
    provider as pr, make_pvs, parent_doc, get_cell_type, Target,
    FrameSearchFlag, date_to_float, float_to_date)
from py4lo_typing import (UnoCell, UnoSheet, UnoSpreadsheetDocument,
                          StrPath, UnoService, UnoRange)

try:
    # noinspection PyUnresolvedReferences
    from com.sun.star.lang import Locale


    class NumberFormat:
        # noinspection PyUnresolvedReferences
        from com.sun.star.util.NumberFormat import (DATE, TIME, DATETIME,
                                                    LOGICAL)
except (ModuleNotFoundError, ImportError):
    from _mock_constants import (Locale, NumberFormat)


class CellTyping(Enum):
    """
    The typing for the cell read (see. create_read_cell).

    * "String" means: always return a string
    * "Minimal" means: return None, a string or a float
    * "Accurate" means: return None, a string, a float, a date or a bool
    """
    String = 0
    Minimal = 1
    Accurate = 2


##########
# Reader #
##########

def create_read_cell(cell_typing: CellTyping = CellTyping.Minimal,
                     oFormats: Optional[UnoService] = None
                     ) -> Callable[[UnoCell], Any]:
    """
    Create a function to read a cell.

    @param cell_typing: one of `CellTyping.String` (return the String value),
                      `CellTyping.Minimal` (String or Value),
                      `CellTyping.Accurate` (the most accurate type)
    @param oFormats: the container for NumberFormats.
    @return: a function to read the cell value
    """

    def read_cell_string(oCell: UnoCell) -> str:
        """
        Read a cell value as a string

        @param oCell: the cell
        @return: the cell value as string
        """
        return oCell.String

    def read_cell_minimal(oCell: UnoCell) -> Any:
        """
        Read a cell value as None, a float or a string (as in a DataArray)

        @param oCell: the cell
        @return: the cell value as float or str
        """
        cell_type = get_cell_type(oCell)

        if cell_type == 'EMPTY':
            return None
        elif cell_type == 'TEXT':
            return oCell.String
        elif cell_type == 'VALUE':
            return oCell.Value
        else:
            raise ValueError()

    def read_cell_accurate(oCell: UnoCell) -> Any:
        """
        Read a cell value as any value.

        @param oCell: the cell
        @return: the cell value as float, bool, date or str
        """
        cell_type = get_cell_type(oCell)

        if cell_type == 'EMPTY':
            return None
        elif cell_type == 'TEXT':
            return oCell.String
        elif cell_type == 'VALUE':
            key = oCell.NumberFormat
            cell_data_type = oFormats.getByKey(key).Type
            value = oCell.Value
            if cell_data_type in {NumberFormat.DATE, NumberFormat.DATETIME,
                                  NumberFormat.TIME}:
                return float_to_date(value)
            elif cell_data_type == NumberFormat.LOGICAL:
                return bool(value)
            else:
                return value
        else:
            raise ValueError()

    if cell_typing == CellTyping.String:
        return read_cell_string
    elif cell_typing == CellTyping.Minimal:
        return read_cell_minimal
    elif cell_typing == CellTyping.Accurate:
        if oFormats is None:
            raise ValueError("Need formats to type all values")
        return read_cell_accurate
    else:
        raise ValueError("cell_typing must be one of TYPE_* values")


class reader(Iterator[List[Any]]):
    """
    A reader that returns rows of a range as lists of values.

    Example:
    ```
    oRange = ...
    r = reader.from_typing(oRange, cell_typing=CellTyping.Accurate)
    for row in r:
        print(row)
    ```
    """

    @staticmethod
    def from_typing(
            oRange: UnoRange,
            cell_typing: CellTyping = CellTyping.Minimal,
            oFormats: Optional[UnoService] = None
    ) -> "reader":
        if cell_typing == CellTyping.Minimal:
            return iter(oRange.DataArray)

        oSheet = oRange.Spreadsheet
        if cell_typing == CellTyping.Accurate and oFormats is None:
            oFormats = oSheet.DrawPage.Forms.Parent.NumberFormats

        read_cell = create_read_cell(cell_typing, oFormats)
        return reader(oRange, read_cell)

    @staticmethod
    def from_read_cell(
            oRange: UnoRange, read_cell: Callable[[UnoCell], Any]
    ) -> "reader":
        """
        @param oRange: the range
        @param read_cell: the read_cell function
        """
        return reader(oRange, read_cell)

    def __init__(self, oRange: UnoRange,
                 read_cell: Callable[[UnoCell], Any]):
        """
        @param oRange: the range
        @param read_cell: the read_cell function
        """
        self._read_cell = read_cell
        self._oSheet = oRange.Spreadsheet
        self.line_num = 0
        self._oRangeAddress = oRange.RangeAddress

    def __iter__(self) -> "reader":
        return self

    def __next__(self) -> List[Any]:
        i = self._oRangeAddress.StartRow + self.line_num
        if i > self._oRangeAddress.EndRow:
            raise StopIteration

        self.line_num += 1
        row = [self._read_cell(self._oSheet.getCellByPosition(j, i))
               for j in range(self._oRangeAddress.StartColumn,
                              self._oRangeAddress.EndColumn + 1)]

        # left strip the row
        i = len(row) - 1
        while row[i] is None and i > 0:
            i -= 1
        return row[:i + 1]


class dict_reader:
    """
    A reader that returns rows of a range as dicts.

    Example:
    ```
    oRange = ...
    r = dict_reader.from_typing(oRange, cell_typing=CellTyping.Accurate)
    for d in r:
        print(d)
    ```
    """

    @staticmethod
    def from_typing(
            oRange: UnoRange,
            fieldnames: Optional[Tuple[str, ...]] = None,
            restkey: Optional[str] = None,
            restval: Optional[Any] = None,
            cell_typing: CellTyping = CellTyping.Minimal,
            oFormats: Optional[UnoService] = None
    ) -> "dict_reader":
        """
        @param oSheet: the sheet
        @param fieldnames: the names of the fields. If None, the fields are
        read from the first line
        @param restkey: the key for additional values
        @param restval: the value for missings fieldnames
        @param cell_typing: the cell typing level (ignored if read_cell is set)
        @param oFormats: the formats of the document (ignored if read_cell is set)
        """
        r = reader.from_typing(oRange, cell_typing, oFormats)
        return dict_reader(r, fieldnames, restkey, restval)

    @staticmethod
    def from_read_cell(
            oRange: UnoRange,
            read_cell: Optional[Callable[[UnoCell], Any]],
            fieldnames: Optional[Tuple[str, ...]] = None,
            restkey: Optional[str] = None,
            restval: Optional[Any] = None,
    ) -> "dict_reader":
        """
        @param oSheet: the sheet
        @param fieldnames: the names of the fields. If None, the fields are
        read from the first line
        @param restkey: the key for additional values
        @param restval: the value for missings fieldnames
        @param read_cell: the read_cell function
        """
        r = reader.from_read_cell(oRange, read_cell)
        return dict_reader(r, fieldnames, restkey, restval)

    def __init__(
            self, r: reader,
            fieldnames: Optional[Tuple[str, ...]] = None,
            restkey: Optional[str] = None,
            restval: Optional[Any] = None
    ):
        """
        @param r: the reader (see. reader)
        @param fieldnames: the names of the fields. If None, the fields are
        read from the first line
        @param cell_typing: the cell typing level (ignored if read_cell is set)
        @param oFormats: the formats of the document (ignored if read_cell is set)
        """
        self._reader = r
        if fieldnames is None:
            self.fieldnames = tuple(
                [v.strip() if isinstance(v, str) else "" for v in
                 next(self._reader)]
            )
        else:
            self.fieldnames = fieldnames
        self._width = len(self.fieldnames)
        self.restkey = restkey
        self.restval = restval

    def __iter__(self):
        return self

    def __next__(self):
        row = next(self._reader)
        row_width = len(row)
        if row_width == self._width:
            return dict(zip(self.fieldnames, row))
        elif row_width < self._width:
            row += [self.restval] * (self._width - row_width)
            return dict(zip(self.fieldnames, row))
        elif self.restkey is None:
            return dict(zip(self.fieldnames, row))
        else:
            d = dict(zip(self.fieldnames, row))
            d[self.restkey] = row[self._width:]
            return d

    @property
    def line_num(self) -> int:
        """The number of the line"""
        return self._reader.line_num


##########
# Writer #
##########
def find_number_format_style(oFormats: UnoService, format_id: NumberFormat,
                             oLocale: Locale = Locale()) -> int:
    """
    Find the number format style number. Use the return value for
    oCell.NumberFormat.

    Example:
    ```
    oCell.NumberFormat = find_number_format_style(
        oDoc.NumberFormats, NumberFormat.DATE)
    ```

    @param oFormats: the formats (see. com.sun.star.util.NumberFormats)
    @param format_id: a NumberFormat
    @param oLocale: the locale (see. com.sun.star.lang.Locale)
    @return: the id of the format
    """
    return oFormats.getStandardFormat(format_id, oLocale)


def create_write_cell(cell_typing: CellTyping = CellTyping.Minimal,
                      oFormats: Optional[UnoService] = None
                      ) -> Callable[[UnoCell, Any], None]:
    """
    Create a cell writer.

    Example:
    ```
    write_cell = create_write_cell(CellTyping.Accurate)
    write_cell(oCell, dt.datetime.now())
    ```

    @param cell_typing: see `create_read_cell`
    @param oFormats: the NumberFormats
    @return: a function
    """

    def write_cell_string(oCell: UnoCell, value: Any):
        """
        Write a cell value.

        @param oCell: the cell
        @param value: the value
        """
        oCell.String = str(value)

    def write_cell_minimal(oCell: UnoCell, value: Any):
        """
        Write a cell value.

        @param oCell: the cell
        @param value: the value
        """
        if value is None:
            oCell.String = ""
        elif isinstance(value, str):
            oCell.String = value
        elif isinstance(value, (date, datetime, time)):
            oCell.Value = date_to_float(value)  # to use oDoc.NullDate
        elif isinstance(value, bool):
            oCell.Value = int(value)
        else:
            oCell.Value = value

    def create_write_cell_all(oFormats: UnoService
                              ) -> Callable[[UnoCell, Any], None]:
        """
        @param oFormats: the formats (see. com.sun.star.util.NumberFormats)
        @return: the function
        """
        # TODO: Add locale parameter
        date_id = find_number_format_style(oFormats, NumberFormat.DATE)
        datetime_id = find_number_format_style(oFormats, NumberFormat.DATETIME)
        boolean_id = find_number_format_style(oFormats, NumberFormat.LOGICAL)

        def write_cell_all(oCell: UnoCell, value: Any):
            """
            Write a cell value and the format according to the type of the value.

            @param oCell: the cell
            @param value: the value
            """
            if value is None:
                oCell.String = ""
            elif isinstance(value, str):
                oCell.String = value
            elif isinstance(value, (datetime, time)):
                oCell.Value = date_to_float(value)  # to use oDoc.NullDate
                oCell.NumberFormat = datetime_id
            elif isinstance(value, date):
                oCell.Value = date_to_float(value)  # to use oDoc.NullDate
                oCell.NumberFormat = date_id
            elif isinstance(value, bool):
                oCell.Value = int(value)
                oCell.NumberFormat = boolean_id
            else:
                oCell.Value = value

        return write_cell_all

    if cell_typing == CellTyping.String:
        return write_cell_string
    elif cell_typing == CellTyping.Minimal:
        return write_cell_minimal
    elif cell_typing == CellTyping.Accurate:
        if oFormats is None:
            raise ValueError("Need formats to type all values")
        return create_write_cell_all(oFormats)
    else:
        raise ValueError("cell_typing must be one of TYPE_* values")


class writer:
    """
    A writer that takes lists of values.

    Example:
    ```
    oSheet = ...
    w = writer.from_typing(oSheet, CellTyping.Accurate)
    w.write_row(["foo", 1.0, dt.datetime.now()])
    ```
    """

    @staticmethod
    def from_typing(
            oSheet: UnoSheet,
            cell_typing: CellTyping = CellTyping.Minimal,
            oFormats: Optional[UnoService] = None,
            initial_pos: Tuple[int, int] = (0, 0)) -> "writer":
        """
        @param oSheet: the destination sheet
        @param cell_typing: a cell typing level
        @param oFormats: the formats
        @param initial_pos: the initial position on the sheet, as a tuple
        """
        if cell_typing == CellTyping.Accurate and oFormats is None:
            oFormats = oSheet.DrawPage.Forms.Parent.NumberFormats
        write_cell = create_write_cell(cell_typing, oFormats)
        return writer(oSheet, write_cell, initial_pos)

    @staticmethod
    def from_write_cell(oSheet: UnoSheet,
                        write_cell: Callable[[UnoCell, Any], None],
                        initial_pos: Tuple[int, int] = (0, 0)) -> "writer":
        """
        @param oSheet: the destination sheet
        @param write_cell: the write_cell function (write a value in a cell)
        @param initial_pos: the initial position on the sheet, as a tuple
        """
        return writer(oSheet, write_cell, initial_pos)

    def __init__(self, oSheet: UnoSheet,
                 write_cell: Callable[[UnoCell, Any], None],
                 initial_pos: Tuple[int, int] = (0, 0)):
        """
        @param oSheet: the destination sheet
        @param cell_typing: a cell typing level
        @param oFormats: the formats
        @param write_cell: the write_cell function (write a value in a cell)
        @param initial_pos: the initial position on the sheet, as a tuple
        """
        self._oSheet = oSheet
        self._write_cell = write_cell
        self._row, self._base_col = initial_pos

    def writerow(self, row: List[Any]):
        """
        Write a
        @param row:
        @return:
        """
        for i, value in enumerate(row, self._base_col):
            oCell = self._oSheet.getCellByPosition(i, self._row)
            self._write_cell(oCell, value)
        self._row += 1

    def writerows(self, rows):
        for row in rows:
            self.writerow(row)


class dict_writer:
    """
    A writer that takes dicts

    Example:
    ```
    oSheet = ...
    w = dict_writer.from_typing(oSheet, CellTyping.Accurate)
    w.write_row({"x": "foo", "y": 1.0, "z": dt.datetime.now()})
    ```
    """

    @staticmethod
    def from_typing(
            oSheet: UnoSheet,
            fieldnames: List[str],
            cell_typing: CellTyping = CellTyping.Minimal,
            oFormats: Optional[UnoService] = None,
            restval: str = '', extrasaction: str = 'raise',
            initial_pos: Tuple[int, int] = (0, 0)) -> "dict_writer":
        """
        @param oSheet: the destination sheet
        @param fieldnames: the names of the fields. If None, the fields are
        read from the first line
        @param restval: the value for missings fieldnames
        @param extrasaction: if "raise", raise an exception
        @param cell_typing: a cell typing level
        @param oFormats: the formats
        @param initial_pos: the initial position on the sheet, as a tuple
        """
        if cell_typing == CellTyping.Accurate and oFormats is None:
            oFormats = oSheet.DrawPage.Forms.Parent.NumberFormats
        write_cell = create_write_cell(cell_typing, oFormats)
        return dict_writer(oSheet, write_cell, fieldnames, restval,
                           extrasaction, initial_pos)

    @staticmethod
    def from_write_cell(
            oSheet: UnoSheet,
            fieldnames: List[str],
            write_cell: Callable[[UnoCell, Any], None],
            restval: str = '', extrasaction: str = 'raise',
            initial_pos: Tuple[int, int] = (0, 0)) -> "dict_writer":
        """
        @param oSheet: the destination sheet
        @param fieldnames: the names of the fields. If None, the fields are
        read from the first line
        @param restval: the value for missings fieldnames
        @param extrasaction: if "raise", raise an exception
        @param write_cell: the write_cell function (write a value in a cell)
        @param initial_pos: the initial position on the sheet, as a tuple
        """
        return dict_writer(oSheet, write_cell, fieldnames, restval,
                           extrasaction, initial_pos)

    def __init__(
            self, oSheet: UnoSheet,
            write_cell: Callable[[UnoCell, Any], None],
            fieldnames: List[str],
            restval: str = '', extrasaction: str = 'raise',
            initial_pos: Tuple[int, int] = (0, 0)
    ):
        """
        @param oSheet: the destination sheet
        @param fieldnames: the names of the fields. If None, the fields are
        read from the first line
        @param restval: the value for missings fieldnames
        @param extrasaction: if "raise", raise an exception
        @param write_cell: the write_cell function (write a value in a cell)
        @param initial_pos: the initial position on the sheet, as a tuple
        """
        self.writer = writer.from_write_cell(oSheet, write_cell, initial_pos)
        self.fieldnames = fieldnames
        self._set_fieldnames = set(fieldnames)
        self.restval = restval
        self.extrasaction = extrasaction

    def writeheader(self):
        """
        Write the header
        """
        self.writer.writerow(self.fieldnames)

    def writerow(self, row: Mapping[str, Any]):
        """
        Write a row

        @param row: a mapping key -> value.
        """
        if self.extrasaction == 'raise' and set(row) - self._set_fieldnames:
            raise ValueError()
        flat_row = [row.get(name, self.restval) for name in self.fieldnames]
        self.writer.writerow(flat_row)

    def writerows(self, rows: Iterable[Mapping[str, Any]]):
        """
        Write rows

        @param rows: an iterable of mappings key -> value.
        """
        for row in rows:
            self.writerow(row)


#####################
# Import/Export CSV #
#####################
#
# A CSV import requires a FilterName and some FilterOptions:
#
###
# FilterName
# See https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html
###
class Filter(str, Enum):
    """
    See. https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html
    """
    XML = "StarOffice XML (Calc)"  # Standard XML filter
    # XML filter for templates
    XML_TEMPLATE = "calc_StarOffice_XML_Calc_Template"
    STARCALC_5 = "StarCalc 5.0"  # The binary format of StarOffice Calc 5.x
    # StarOffice Calc 5.x templates
    STARCALC_5_TEMPLATE = "StarCalc 5.0 Vorlage/Template"
    STARCALC_4 = "StarCalc 4.0"  # The binary format of StarCalc 4.x
    # StarCalc 4.x templates
    STARCALC_4_TEMPLATE = "StarCalc 4.0 Vorlage/Template"
    STARCALC_3 = "StarCalc 3.0"  # The binary format of StarCalc 3.x
    # StarCalc 3.x templates
    STARCALC_3_TEMPLATE = "StarCalc 3.0 Vorlage/Template"
    HTML = "HTML (StarCalc)"  # HTML filter
    # HTML filter for external data queries
    HTML_WEBQUERY = "calc_HTML_WebQuery"
    EXCEL_97 = "MS Excel 97"  # Microsoft Excel 97/2000/XP
    # Microsoft Excel 97/2000/XP templates
    EXCEL_97_TEMPLATE = "MS Excel 97 Vorlage/Template"
    EXCEL_95 = "MS Excel 95"  # Microsoft Excel 5.0/95
    # Microsoft Excel 5.0/95 templates
    EXCEL_95_TEMPLATE = "MS Excel 95 Vorlage/Template"
    EXCEL_2_3_4 = "MS Excel 4.0"  # Microsoft Excel 2.1/3.0/4.0
    # Microsoft Excel 2.1/3.0/4.0 templates
    EXCEL_2_3_4_TEMPLATE = "MS Excel 4.0 Vorlage/Template"
    LOTUS = "Lotus"  # Lotus 1-2-3
    CSV = "Text - txt - csv (StarCalc)"  # Comma separated values
    RTF = "Rich Text Format (StarCalc)"  #
    DBASE = "dBase"  # dBase
    SYLK = "SYLK"  # Symbolic Link
    DIF = "DIF"  # Data Interchange Format


###
# FilterOptipns
# See https://help.libreoffice.org/latest/en-US/text/shared/guide/csv_params.html
# token 1 & 2 : field separator & delimiter
# token 3 : char set
# token 4 : line start
# token 5 : cell formats
# token 6 : language id
# ...
###
##########
# TOKEN 3
# see https://api.libreoffice.org/docs/cpp/ref/a00391_source.html (rtl/textenc.h)
##########
CHARSET_ID_BY_NAME = {
    'unknown': 0,
    'cp1252': 1,
    'mac_roman': 2,
    'cp437': 3,
    'cp850': 4,
    'cp860': 5,
    'cp861': 6,
    'cp863': 7,
    'cp865': 8,
    sys.getdefaultencoding(): 9,
    'symbol': 10,
    'ascii': 11,
    'iso8859_1': 12,
    'iso8859_2': 13,
    'iso8859_3': 14,
    'iso8859_4': 15,
    'iso8859_5': 16,
    'iso8859_6': 17,
    'iso8859_7': 18,
    'iso8859_8': 19,
    'iso8859_9': 20,
    'iso8859_14': 21,
    'iso8859_15': 22,
    'cp737': 23,
    'cp775': 24,
    'cp852': 25,
    'cp855': 26,
    'cp857': 27,
    'cp862': 28,
    'cp864': 29,
    'cp866': 30,
    'cp869': 31,
    'cp874': 32,
    'cp1250': 33,
    'cp1251': 34,
    'cp1253': 35,
    'cp1254': 36,
    'cp1255': 37,
    'cp1256': 38,
    'cp1257': 39,
    'cp1258': 40,
    'mac_arabic': 41,
    'mac_centeuro': 42,
    'mac_croatian': 43,
    'mac_cyrillic': 44,
    'mac_devanagari': 45,
    'mac_farsi': 46,
    'mac_greek': 47,
    'mac_gujarati': 48,
    'mac_gurmukhi': 49,
    'mac_hebrew': 50,
    'mac_iceland': 51,
    'mac_romanian': 52,
    'mac_thai': 53,
    'mac_turkish': 54,
    'mac_ukrainian': 55,
    'mac_chinsimp': 56,
    'mac_chintrad': 57,
    'mac_japanese': 58,
    'mac_korean': 59,
    'cp932': 60,
    'cp936': 61,
    'cp949': 62,
    'cp950': 63,
    'shift_jis': 64,
    'gb2312': 65,
    'gbt12345': 66,
    'gbk': 67,
    'big5': 68,
    'euc_jp': 69,
    'euc_cn': 70,
    'euc_tw': 71,
    'iso2022_jp': 72,
    'iso2022_cn': 73,
    'koi8_r': 74,
    'utf_7': 75,
    'utf_8': 76,
    'iso8859_10': 77,
    'iso8859_13': 78,
    'euc_kr': 79,
    'iso2022_kr': 80,
    'jis_x_0201': 81,
    'jis_x_0208': 82,
    'jis_x_0212': 83,
    'cp1361': 84,
    'gb18030': 85,
    'big5hkscs': 86,
    'tis_620': 87,
    'koi8_u': 88,
    'iscii_devanagari': 89,
    'java_utf8': 90,
    'adobe_standard': 91,
    'adobe_symbol': 92,
    'ptcp154': 93,
    'adobe_dingbats': 94,
    'user_start': 32768,
    'user_end': 61439,
    'utf_32': 65534,
    'utf_16': 65535,
}
"""
A mapping charset name -> charset id for token 3 of CSV FilterOptions 
(see https://api.libreoffice.org/docs/cpp/ref/a00391_source.html (rtl/textenc.h)
"""


##########
# TOKEN 5
# See.https://help.libreoffice.org/latest/en-US/text/shared/guide/csv_params.html
##########
class Format(IntEnum):
    """
    Formatting Codes for Token 5

    See. https://help.libreoffice.org/latest/en-US/text/shared/guide/csv_params.html
    "Formatting Codes for Token 5".
    """
    STANDARD = 1  # Standard
    TEXT = 2  # Text
    MM_DD_YY = 3  # MM/DD/YY
    DD_MM_YY = 4  # DD/MM/YY
    YY_MM_DD = 5  # YY/MM/DD
    IGNORE = 9  # IGNORE FIELD (do not import)
    US = 10  # US-English


##########
# TOKEN 6
# see https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6c085406-a698-4e12-9d4d-c3b0ee3dbc4a
# and https://help.libreoffice.org/latest/en-US/text/shared/guide/csv_params.html,
##########
LANGUAGE_ID_BY_CODE = {
    "ar_SA": 1025,
    "bg_BG": 1026,
    "ca_ES": 1027,
    "zh_TW": 1028,
    "cs_CZ": 1029,
    "da_DK": 1030,
    "de_DE": 1031,
    "el_GR": 1032,
    "en_US": 1033,
    "es_ES": 1034,
    "fi_FI": 1035,
    "fr_FR": 1036,
    "he_IL": 1037,
    "hu_HU": 1038,
    "is_IS": 1039,
    "it_IT": 1040,
    "ja_JP": 1041,
    "ko_KR": 1042,
    "nl_NL": 1043,
    "nb_NO": 1044,
    "pl_PL": 1045,
    "pt_BR": 1046,
    "rm_CH": 1047,
    "ro_RO": 1048,
    "ru_RU": 1049,
    "hr_HR": 1050,
    "sk_SK": 1051,
    "sq_AL": 1052,
    "sv_SE": 1053,
    "th_TH": 1054,
    "tr_TR": 1055,
    "ur_PK": 1056,
    "id_ID": 1057,
    "uk_UA": 1058,
    "be_BY": 1059,
    "sl_SI": 1060,
    "et_EE": 1061,
    "lv_LV": 1062,
    "lt_LT": 1063,
    "tg_Cyrl_TJ": 1064,
    "fa_IR": 1065,
    "vi_VN": 1066,
    "hy_AM": 1067,
    "az_Latn_AZ": 1068,
    "eu_ES": 1069,
    "wen_DE": 1070,
    "mk_MK": 1071,
    "st_ZA": 1072,
    "ts_ZA": 1073,
    "tn_ZA": 1074,
    "ven_ZA": 1075,
    "xh_ZA": 1076,
    "zu_ZA": 1077,
    "af_ZA": 1078,
    "ka_GE": 1079,
    "fo_FO": 1080,
    "hi_IN": 1081,
    "mt_MT": 1082,
    "se_NO": 1083,
    "gd_GB": 1084,
    "yi": 1085,
    "ms_MY": 1086,
    "kk_KZ": 1087,
    "ky_KG": 1088,
    "sw_KE": 1089,
    "tk_TM": 1090,
    "uz_Latn_UZ": 1091,
    "tt_RU": 1092,
    "bn_IN": 1093,
    "pa_IN": 1094,
    "gu_IN": 1095,
    "or_IN": 1096,
    "ta_IN": 1097,
    "te_IN": 1098,
    "kn_IN": 1099,
    "ml_IN": 1100,
    "as_IN": 1101,
    "mr_IN": 1102,
    "sa_IN": 1103,
    "mn_MN": 1104,
    "bo_CN": 1105,
    "cy_GB": 1106,
    "km_KH": 1107,
    "lo_LA": 1108,
    "my_MM": 1109,
    "gl_ES": 1110,
    "kok_IN": 1111,
    "mni": 1112,
    "sd_IN": 1113,
    "syr_SY": 1114,
    "si_LK": 1115,
    "chr_US": 1116,
    "iu_Cans_CA": 1117,
    "am_ET": 1118,
    "tmz": 1119,
    "ks_Arab_IN": 1120,
    "ne_NP": 1121,
    "fy_NL": 1122,
    "ps_AF": 1123,
    "fil_PH": 1124,
    "dv_MV": 1125,
    "bin_NG": 1126,
    "fuv_NG": 1127,
    "ha_Latn_NG": 1128,
    "ibb_NG": 1129,
    "yo_NG": 1130,
    "quz_BO": 1131,
    "nso_ZA": 1132,
    "ig_NG": 1136,
    "kr_NG": 1137,
    "gaz_ET": 1138,
    "ti_ER": 1139,
    "gn_PY": 1140,
    "haw_US": 1141,
    "la": 1142,
    "so_SO": 1143,
    "ii_CN": 1144,
    "pap_AN": 1145,
    "ug_Arab_CN": 1152,
    "mi_NZ": 1153,
    "ar_IQ": 2049,
    "zh_CN": 2052,
    "de_CH": 2055,
    "en_GB": 2057,
    "es_MX": 2058,
    "fr_BE": 2060,
    "it_CH": 2064,
    "nl_BE": 2067,
    "nn_NO": 2068,
    "pt_PT": 2070,
    "ro_MD": 2072,
    "ru_MD": 2073,
    "sr_Latn_CS": 2074,
    "sv_FI": 2077,
    "ur_IN": 2080,
    "az_Cyrl_AZ": 2092,
    "ga_IE": 2108,
    "ms_BN": 2110,
    "uz_Cyrl_UZ": 2115,
    "bn_BD": 2117,
    "pa_PK": 2118,
    "mn_Mong_CN": 2128,
    "bo_BT": 2129,
    "sd_PK": 2137,
    "tzm_Latn_DZ": 2143,
    "ks_Deva_IN": 2144,
    "ne_IN": 2145,
    "quz_EC": 2155,
    "ti_ET": 2163,
    "ar_EG": 3073,
    "zh_HK": 3076,
    "de_AT": 3079,
    "en_AU": 3081,
    "fr_CA": 3084,
    "sr_Cyrl_CS": 3098,
    "quz_PE": 3179,
    "ar_LY": 4097,
    "zh_SG": 4100,
    "de_LU": 4103,
    "en_CA": 4105,
    "es_GT": 4106,
    "fr_CH": 4108,
    "hr_BA": 4122,
    "ar_DZ": 5121,
    "zh_MO": 5124,
    "de_LI": 5127,
    "en_NZ": 5129,
    "es_CR": 5130,
    "fr_LU": 5132,
    "bs_Latn_BA": 5146,
    "ar_MO": 6145,
    "en_IE": 6153,
    "es_PA": 6154,
    "fr_MC": 6156,
    "ar_TN": 7169,
    "en_ZA": 7177,
    "es_DO": 7178,
    "fr_029": 7180,
    "ar_OM": 8193,
    "en_JM": 8201,
    "es_VE": 8202,
    "fr_RE": 8204,
    "ar_YE": 9217,
    "en_029": 9225,
    "es_CO": 9226,
    "fr_CG": 9228,
    "ar_SY": 10241,
    "en_BZ": 10249,
    "es_PE": 10250,
    "fr_SN": 10252,
    "ar_JO": 11265,
    "en_TT": 11273,
    "es_AR": 11274,
    "fr_CM": 11276,
    "ar_LB": 12289,
    "en_ZW": 12297,
    "es_EC": 12298,
    "fr_CI": 12300,
    "ar_KW": 13313,
    "en_PH": 13321,
    "es_CL": 13322,
    "fr_ML": 13324,
    "ar_AE": 14337,
    "en_ID": 14345,
    "es_UY": 14346,
    "fr_MA": 14348,
    "ar_BH": 15361,
    "en_HK": 15369,
    "es_PY": 15370,
    "fr_HT": 15372,
    "ar_QA": 16385,
    "en_IN": 16393,
    "es_BO": 16394,
    "en_MY": 17417,
    "es_SV": 17418,
    "en_SG": 18441,
    "es_HN": 18442,
    "es_NI": 19466,
    "es_PR": 20490,
    "es_US": 21514,
    "es_419": 58378,
    "fr_015": 58380
}
"""
A mapping language_territory -> locale id for token 5 of CSV FilterOptions 
"""


def import_from_csv(oDoc: UnoSpreadsheetDocument, sheet_name: str,
                    dest_position: int,
                    path: StrPath, *args, **kwargs):
    """
    @param oDoc: the document
    @param sheet_name: the target sheet name
    @param dest_position: the target sheet position
    @param path: path to the file
    @param args: if one arg, it is the Python csv dialect, else
    @param encoding: the source file encoding
    @param first_line: the first line
    @param format_by_idx: a mappin col -> type (type is FORMAT_MM_DD_YY)
    @param language_code: en_US
    @param detect_special_numbers: if true, detect numbers
    """
    filter_options = create_import_filter_options(*args, **kwargs)
    pvs = make_pvs({"FilterName": Filter.CSV, "FilterOptions": filter_options,
                    "Hidden": True})

    url = uno_path_to_url(path)
    oDesktop = pr.desktop

    oDoc.lockControllers()
    try:
        oSource = cast(
            UnoSpreadsheetDocument,
            oDesktop.loadComponentFromURL(
                url, Target.BLANK, FrameSearchFlag.AUTO, pvs)
        )
        try:
            oSource.Sheets.getByIndex(0).Name = sheet_name
            name = oSource.Sheets.ElementNames[0]
            oDoc.Sheets.importSheet(oSource, name, dest_position)
        finally:
            oSource.close(True)
    finally:
        oDoc.unlockControllers()


# IMPORT

def create_import_filter_options(*args, **kwargs) -> str:
    """
    Create FilterOptions string for CSV import.
    See https://help.libreoffice.org/latest/en-US/text/shared/guide/csv_params.html

    Takes 0 or 1 positional argument.
    If there is 1 positional argument, this is the Python csv dialect.
    If there is no positional argument, then named arguments are the definition
    of the CSV FilterOptions.

    Example 1:
    ```
    fo = create_import_filter_options(csv.excel_tab, first_line=5)
    ```

    Example 2:
    ```
    fo = create_import_filter_options(delimiter=",", quotechar='"', first_line=5)
    ```

    @param args: the Python csv dialect or nothing
    @param kwargs: delimiter, quotechar, quoted_field_as_text, encoding,
        language_code, first_line, format_by_idx, detect_special_numbers are
        allowed
    """
    if len(args) == 1:
        dialect = args[0]
        delimiter = kwargs.pop("delimiter", dialect.delimiter)
        quotechar = kwargs.pop("quotechar", dialect.quotechar)
        quoted_field_as_text = kwargs.pop("quoted_field_as_text",
                                          dialect.quoting == csv.QUOTE_ALL)
        return _create_import_filter_options(
            delimiter=delimiter, quotechar=quotechar,
            quoted_field_as_text=quoted_field_as_text,
            **kwargs
        )
    elif len(args) == 0:
        return _create_import_filter_options(**kwargs)
    else:
        raise ValueError("At most one positional parameter allowed")


def _create_import_filter_options(
        delimiter: str = ",", quotechar: str = '"',
        quoted_field_as_text: bool = False,
        encoding: str = "utf-8", language_code=locale.getlocale()[0],
        first_line: int = 1,
        format_by_idx: Optional[Mapping[int, Format]] = None,
        detect_special_numbers: bool = False) -> str:
    """
    Create FilterOptions string for CSV import.
    See https://help.libreoffice.org/latest/en-US/text/shared/guide/csv_params.html

    @param delimiter: the delimiter
    @param quotechar: the quotechar
    @param quoted_field_as_text: see checkbox
    @param encoding: the source file encoding
    @param language_code: en_US
    @param first_line: the first line
    @param format_by_idx: a mappin col -> type (type is FORMAT_MM_DD_YY)
    @param detect_special_numbers: if true, detect numbers
    @return: a filter options string
    """
    quoting = "true" if quoted_field_as_text else "false"
    detect = "true" if detect_special_numbers else "false"
    tokens = _base_filter_tokens(
        delimiter, quotechar, encoding, language_code, first_line,
        format_by_idx
    ) + [quoting, detect]
    return ",".join(tokens)


def _base_filter_tokens(
        delimiter: str, quotechar: str, encoding: str, language_code: str,
        first_line: int, format_by_idx: Optional[Mapping[int, int]]
) -> List[str]:
    """
    The base parameters, common to import/export
    See: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options

    @param delimiter: the delimiter
    @param quotechar: the quotechar
    @param encoding: the encoding
    @param first_line: the first line
    @param format_by_idx: a mapping field index (starting at 1) -> field format
    @return: a list of options
    """
    encoding_index = _get_charset_id(encoding)
    code = "" if language_code is None else language_code
    language_id = LANGUAGE_ID_BY_CODE.get(code, 0)
    field_formats = _build_field_formats(format_by_idx)
    return [str(ord(delimiter)), str(ord(quotechar)), str(encoding_index),
            str(first_line), field_formats, str(language_id)]


# EXPORT
def export_to_csv(oSheet: UnoSheet, path: StrPath, *args, **kwargs):
    """
    Save a sheet to a csv file.

    @param oSheet: the sheet
    @param path: the path of the target file
    @param args: see create_export_filter_options
    @param kwargs: see create_export_filter_options
    """
    overwrite = kwargs.pop("overwrite", True)
    filter_options = create_export_filter_options(*args, **kwargs)
    pvs = make_pvs({"FilterName": Filter.CSV, "FilterOptions": filter_options,
                    "Overwrite": overwrite})
    oDoc = parent_doc(oSheet)
    oActive = oDoc.CurrentController.ActiveSheet
    oDoc.lockControllers()
    try:
        oDoc.CurrentController.ActiveSheet = oSheet
        url = uno_path_to_url(path)
        oDoc.storeToURL(url, pvs)
        oDoc.CurrentController.ActiveSheet = oActive
    finally:
        oDoc.unlockControllers()


def _get_charset_id(encoding: str) -> int:
    """
    Takes the encoding as a string and return the charset id.

    @param encoding: the encoding
    @return: the charset id
    """
    norm_encoding = encodings.normalize_encoding(encoding)
    norm_encoding = encodings.aliases.aliases.get(
        norm_encoding.lower(), norm_encoding)
    return CHARSET_ID_BY_NAME.get(norm_encoding, 0)


def _build_field_formats(format_by_idx: Optional[Mapping[int, int]]) -> str:
    """
    Build the token 5

    @param format_by_idx: a mapping idx of the column -> format (see Format)
    @return: then token 5 as a string
    """
    if format_by_idx is None:
        field_formats = ""
    else:
        field_formats = "/".join(["{}/{}".format(idx, fmt)
                                  for idx, fmt in format_by_idx.items()])
    return field_formats


def create_export_filter_options(*args, **kwargs) -> str:
    """
    Create the FilterOptions value for a CSV export.

    See https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options

    Example 1:
    ```
    fo = create_export_filter_options(csv.excel_tab, first_line=5)
    ```

    Example 2:
    ```
    fo = create_export_filter_options(delimiter=",", quotechar='"', first_line=5)
    ```

    @param args: the Python csv dialect or nothing
    @param kwargs: delimiter, quotechar, store_numeric_cells_as_text, encoding,
    language_code, first_line, format_by_idx, save_cell_contents_as_shown
    @return: the filter options
    """
    if len(args) == 1:
        dialect = args[0]
        delimiter = kwargs.pop("delimiter", dialect.delimiter)
        quotechar = kwargs.pop("quotechar", dialect.quotechar)
        quote_all_text_cells = kwargs.pop("store_numeric_cells_as_text",
                                          dialect.quoting == csv.QUOTE_ALL)
        return _create_export_filter_options(
            delimiter=delimiter, quotechar=quotechar,
            store_numeric_cells_as_text=quote_all_text_cells,
            **kwargs
        )
    elif len(args) == 0:
        return _create_export_filter_options(**kwargs)
    else:
        raise ValueError("At most one positional parameters allowed")


def _create_export_filter_options(
        delimiter: str = ",", quotechar: str = '"',
        store_numeric_cells_as_text: bool = False,
        encoding: str = "utf-8", language_code=locale.getlocale()[0],
        first_line: int = 1,
        format_by_idx: Optional[Mapping[int, Format]] = None,
        save_cell_contents_as_shown: bool = True) -> str:
    """
    Create FilterOptions string for CSV import.
    See https://help.libreoffice.org/latest/en-US/text/shared/guide/csv_params.html

    @param delimiter: the delimiter
    @param quotechar: the quotechar
    @param store_numeric_cells_as_text: see checkbox
    @param encoding: the source file encoding
    @param language_code: en_US
    @param first_line: the first line
    @param format_by_idx: a mappin col -> type (type is FORMAT_MM_DD_YY)
    @param save_cell_contents_as_shown: if true, detect numbers
    @return: a filter options string
    """
    store_as_text = "true" if store_numeric_cells_as_text else "false"
    save_as_shown = "true" if save_cell_contents_as_shown else "false"
    tokens = _base_filter_tokens(
        delimiter, quotechar, encoding, language_code, first_line,
        format_by_idx
    ) + [store_as_text, store_as_text, save_as_shown]
    return ",".join(tokens)
