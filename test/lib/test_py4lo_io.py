#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2022 J. Férard <https://github.com/jferard>
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
import unittest
from unittest.mock import Mock, patch
import datetime as dt

from py4lo_io import (create_import_filter_options,
                      create_export_filter_options, Format, create_read_cell,
                      CellTyping)


class NumberFormat:
    from com.sun.star.util.NumberFormat import (DATE, TIME, DATETIME, LOGICAL,
                                                NUMBER)


class Py4LOIOTestCase(unittest.TestCase):
    def test_create_read_cell_none(self):
        # prepare
        oCell = Mock(String="foo")

        # play
        rc = create_read_cell(CellTyping.String)

        # verify
        self.assertEqual("foo", rc(oCell))

    @patch("py4lo_io.get_cell_type")
    def test_create_read_cell_minimal_empty(self, gct):
        # prepare
        oCell = Mock()
        gct.side_effect = ['EMPTY']

        # play
        rc = create_read_cell(CellTyping.Minimal)

        # verify
        self.assertIsNone(rc(oCell))

    @patch("py4lo_io.get_cell_type")
    def test_create_read_cell_minimal_text(self, gct):
        # prepare
        oCell = Mock(String="foo")
        gct.side_effect = ['TEXT']

        # play
        rc = create_read_cell(CellTyping.Minimal)

        # verify
        self.assertEqual("foo", rc(oCell))

    @patch("py4lo_io.get_cell_type")
    def test_create_read_cell_minimal_value(self, gct):
        # prepare
        oCell = Mock(Value=3.14)
        gct.side_effect = ['VALUE']

        # play
        rc = create_read_cell(CellTyping.Minimal)

        # verify
        self.assertEqual(3.14, rc(oCell))

    @patch("py4lo_io.get_cell_type")
    def test_create_read_cell_minimal_other(self, gct):
        # prepare
        oCell = Mock()
        gct.side_effect = ['OTHER']

        # play
        rc = create_read_cell(CellTyping.Minimal)

        # verify
        with self.assertRaises(ValueError):
            rc(oCell)

    @patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_empty(self, gct):
        # prepare
        oCell = Mock()
        oFormats = Mock()
        gct.side_effect = ['EMPTY']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        self.assertIsNone(rc(oCell))

    @patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_text(self, gct):
        # prepare
        oCell = Mock(String="foo")
        oFormats = Mock()
        gct.side_effect = ['TEXT']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        self.assertEqual("foo", rc(oCell))

    @patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_date(self, gct):
        # prepare
        oCell = Mock(Value=44000, NumberFormat=3)
        oFormats = Mock()
        oFormats.getByKey.side_effect = [Mock(Type=NumberFormat.DATE)]
        gct.side_effect = ['VALUE']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        self.assertEqual(dt.datetime(2020, 6, 18, 0, 0), rc(oCell))

    @patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_logical(self, gct):
        # prepare
        oCell = Mock(Value=1, NumberFormat=3)
        oFormats = Mock()
        oFormats.getByKey.side_effect = [Mock(Type=NumberFormat.LOGICAL)]
        gct.side_effect = ['VALUE']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        self.assertIs(True, rc(oCell))

    @patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_number(self, gct):
        # prepare
        oCell = Mock(Value=10, NumberFormat=3)
        oFormats = Mock()
        oFormats.getByKey.side_effect = [Mock(Type=NumberFormat.NUMBER)]
        gct.side_effect = ['VALUE']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        self.assertEqual(10, rc(oCell))

    @patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_other(self, gct):
        # prepare
        oCell = Mock()
        oFormats = Mock()
        gct.side_effect = ['']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        with self.assertRaises(ValueError):
            rc(oCell)

    def test_empty_import_options(self):
        self.assertEqual('44,34,76,1,,1033,false,false',
                         create_import_filter_options(language_code="en_US"))

    def test_import_options(self):
        self.assertEqual('44,34,76,1,1/2,1033,false,false',
                         create_import_filter_options(language_code="en_US",
                                                      format_by_idx={
                                                          1: Format.TEXT}))

    def test_empty_export_options(self):
        self.assertEqual('44,34,76,1,,1033,false,false,true',
                         create_export_filter_options(language_code="en_US"))


if __name__ == '__main__':
    unittest.main()
