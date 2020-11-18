#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2019 J. Férard <https://github.com/jferard>
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
from datetime import (date, datetime, time)

from com.sun.star.util import NumberFormat
from com.sun.star.lang import Locale

import py4lo_helper
import py4lo_commons

# values of type_cell
TYPE_NONE = 0
TYPE_MINIMAL = 1
TYPE_ALL = 2


##########
# Reader #
##########

def create_read_cell(type_cell=TYPE_MINIMAL, oFormats=None):
    def read_cell_none(oCell):
        """
        Read a cell value
        @param oCell: the cell
        @return: the cell value as string
        """
        return oCell.String

    def read_cell_minimal(oCell):
        """
        Read a cell value
        @param oCell: the cell
        @return: the cell value as float or str
        """
        cell_type = py4lo_helper.get_cell_type(oCell)

        if cell_type == 'EMPTY':
            return None
        elif cell_type == 'TEXT':
            return oCell.String
        elif cell_type == 'VALUE':
            return oCell.Value

    def read_cell_all(oCell):
        """
        Read a cell value
        @param oCell: the cell
        @return: the cell value as float, bool, date or str
        """
        cell_type = py4lo_helper.get_cell_type(oCell)

        if cell_type == 'EMPTY':
            return None
        elif cell_type == 'TEXT':
            return oCell.String
        elif cell_type == 'VALUE':
            key = oCell.NumberFormat
            cell_data_type = oFormats.getByKey(key).Type
            if cell_data_type in {NumberFormat.DATE, NumberFormat.DATETIME,
                                  NumberFormat.TIME}:
                return py4lo_commons.float_to_date(oCell.Value)
            elif cell_data_type == NumberFormat.LOGICAL:
                return bool(oCell.Value)
            else:
                return oCell.Value

    if type_cell == TYPE_NONE:
        return read_cell_none
    elif type_cell == TYPE_MINIMAL:
        return read_cell_minimal
    elif type_cell == TYPE_ALL:
        if oFormats is None:
            raise ValueError("Need formats to type all values")
        return read_cell_all
    else:
        raise ValueError("type_cell must be one of TYPE_* values")


class reader:
    """
    A reader

    """

    def __init__(self, oSheet, type_cell=TYPE_MINIMAL, oFormats=None,
                 read_cell=None):
        if read_cell is not None:
            self._read_cell = read_cell
        else:
            if type_cell == TYPE_ALL and oFormats is None:
                oFormats = oSheet.DrawPage.Forms.Parent.NumberFormats
            self._read_cell = create_read_cell(type_cell, oFormats)
        self._oSheet = oSheet
        self.line_num = 0
        self._oRangeAddress = py4lo_helper.get_used_range_address(oSheet)

    def __iter__(self):
        return self

    def __next__(self):
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
    def __init__(self, oSheet, fieldnames=None, restkey=None, restval=None,
                 type_cell=TYPE_MINIMAL, oFormats=None, read_cell=None):
        self._reader = reader(oSheet, type_cell, oFormats, read_cell)
        if fieldnames is None:
            self.fieldnames = next(self._reader)
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
    def line_num(self):
        return self._reader.line_num


##########
# Writer #
##########
def find_number_format_style(oFormats, format_id, oLocale=Locale()):
    """
    
    @param oFormats: the formats 
    @param format_id: a NumberFormat
    @param oLocale: the locale
    @return: the id of the format
    """
    return oFormats.getStandardFormat(format_id, oLocale)


def create_write_cell(type_cell=TYPE_MINIMAL, oFormats=None):
    def write_cell_none(oCell, value):
        """
        Write a cell value
        @param oCell: the cell
        @param value: the value
        """
        oCell.String = str(value)

    def write_cell_minimal(oCell, value):
        """
        Write a cell value
        @param oCell: the cell
        @param value: the value
        """
        if value is None:
            oCell.String = ""
        elif isinstance(value, str):
            oCell.String = value
        elif isinstance(value, (date, datetime, time)):
            oCell.Value = py4lo_commons.date_to_float(value)
        else:
            oCell.Value = value

    def create_write_cell_all(oFormats):
        date_id = find_number_format_style(oFormats, NumberFormat.DATE)
        datetime_id = find_number_format_style(oFormats, NumberFormat.DATETIME)
        boolean_id = find_number_format_style(oFormats, NumberFormat.LOGICAL)

        def write_cell_all(oCell, value):
            """
            Write a cell value
            @param oCell: the cell
            @param value: the value
            @return: the cell value as float or text
            """
            if value is None:
                oCell.String = ""
            elif isinstance(value, str):
                oCell.String = value
            elif isinstance(value, (datetime, time)):
                oCell.Value = py4lo_commons.date_to_float(value)
                oCell.NumberFormat = datetime_id
            elif isinstance(value, date):
                oCell.Value = py4lo_commons.date_to_float(value)
                oCell.NumberFormat = date_id
            elif isinstance(value, bool):
                oCell.Value = value
                oCell.NumberFormat = boolean_id
            else:
                oCell.Value = value

        return write_cell_all

    if type_cell == TYPE_NONE:
        return write_cell_none
    elif type_cell == TYPE_MINIMAL:
        return write_cell_minimal
    elif type_cell == TYPE_ALL:
        if oFormats is None:
            raise ValueError("Need formats to type all values")
        return create_write_cell_all(oFormats)
    else:
        raise ValueError("type_cell must be one of TYPE_* values")


class writer:
    """
    A writer
    """

    def __init__(self, oSheet, type_cell=TYPE_MINIMAL, oFormats=None,
                 write_cell=None,
                 initial_pos=(0, 0)):
        self._oSheet = oSheet
        self._row, self._base_col = initial_pos
        if write_cell is not None:
            self._write_cell = write_cell
        else:
            if type_cell == TYPE_ALL and oFormats is None:
                oFormats = oSheet.DrawPage.Forms.Parent.NumberFormats
            self._write_cell = create_write_cell(type_cell, oFormats)

    def writerow(self, row):
        for i, value in enumerate(row, self._base_col):
            oCell = self._oSheet.getCellByPosition(i, self._row)
            self._write_cell(oCell, value)
        self._row += 1

    def writerows(self, rows):
        for row in rows:
            self.writerow(row)


class dict_writer:
    def __init__(self, oSheet, fieldnames, restval='', extrasaction='raise',
                 type_cell=TYPE_MINIMAL, oFormats=None, write_cell=None):
        self.writer = writer(oSheet, type_cell, oFormats, write_cell)
        self.fieldnames = fieldnames
        self._set_fieldnames = set(fieldnames)
        self.restval = restval
        self.extrasaction = extrasaction

    def writeheader(self):
        self.writer.writerow(self.fieldnames)

    def writerow(self, row):
        if self.extrasaction == 'raise' and set(row) - self._set_fieldnames:
            raise ValueError()
        flat_row = [row.get(name, self.restval) for name in self.fieldnames]
        self.writer.writerow(flat_row)

    def writerows(self, rows):
        for row in rows:
            self.writerow(row)