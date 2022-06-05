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
                      CellTyping, reader, dict_reader)


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

    def test_create_read_cell_accurate_wo_formats(self):
        with self.assertRaises(ValueError):
            create_read_cell(CellTyping.Accurate)

    @patch("py4lo_io.get_used_range_address")
    def test_reader_rc(self, gura):
        # prepare
        oSheet = Mock()
        oRangeAddress = Mock(StartColumn=0, StartRow=0, EndColumn=1, EndRow=2)
        gura.side_effect = [oRangeAddress]

        # play
        r = reader(oSheet, read_cell=lambda oCell: 15)

        # verify
        self.assertEqual([
            [15, 15],
            [15, 15],
            [15, 15]
        ], list(r))

    @patch("py4lo_io.get_used_range_address")
    def test_reader(self, gura):
        # prepare
        oSheet = Mock()
        oSheet.getCellByPosition.side_effect = [
            Mock(String="A1"), Mock(String="B1"),
            Mock(String="A2"), Mock(String="B2"),
            Mock(String="A3"), Mock(String="B3"),
        ]
        oRangeAddress = Mock(StartColumn=0, StartRow=0, EndColumn=1, EndRow=2)
        gura.side_effect = [oRangeAddress]

        # play
        r = reader(oSheet, CellTyping.String)

        # verify
        self.assertEqual([
            ['A1', 'B1'], ['A2', 'B2'], ['A3', 'B3']
        ], list(r))

    @patch("py4lo_io.get_cell_type")
    @patch("py4lo_io.get_used_range_address")
    def test_reader_accurate(self, gura, gct):
        # prepare
        gct.side_effect = ['TEXT', 'TEXT', 'TEXT', 'TEXT', 'TEXT', 'EMPTY']
        oSheet = Mock()
        oSheet.getCellByPosition.side_effect = [
            Mock(String="A1"), Mock(String="B1"),
            Mock(String="A2"), Mock(String="B2"),
            Mock(String="A3"), Mock(),
        ]
        oFormats = Mock()
        oRangeAddress = Mock(StartColumn=0, StartRow=0, EndColumn=1, EndRow=2)
        gura.side_effect = [oRangeAddress]

        # play
        r = reader(oSheet, CellTyping.Accurate, oFormats)

        # verify
        self.assertEqual([
            ['A1', 'B1'], ['A2', 'B2'], ['A3']
        ], list(r))

    @patch("py4lo_io.get_used_range_address")
    def test_dict_reader(self, gura):
        # prepare
        oSheet = Mock()
        oSheet.getCellByPosition.side_effect = [
            Mock(String="A1"), Mock(String="B1"),
            Mock(String="A2"), Mock(String="B2"),
            Mock(String="A3"), Mock(String="B3"),
        ]
        oRangeAddress = Mock(StartColumn=0, StartRow=0, EndColumn=1, EndRow=2)
        gura.side_effect = [oRangeAddress]

        # play
        r = dict_reader(oSheet, cell_typing=CellTyping.String)

        # verify
        self.assertEqual([
            {'A1': 'A2', 'B1': 'B2'}, {'A1': 'A3', 'B1': 'B3'}
        ], list(r))

    @patch("py4lo_io.get_used_range_address")
    def test_dict_reader_fieldnames(self, gura):
        # prepare
        oSheet = Mock()
        oSheet.getCellByPosition.side_effect = [
            Mock(String="A1"), Mock(String="B1"),
            Mock(String="A2"), Mock(String="B2"),
            Mock(String="A3"), Mock(String="B3"),
        ]
        oRangeAddress = Mock(StartColumn=0, StartRow=0, EndColumn=1, EndRow=2)
        gura.side_effect = [oRangeAddress]

        # play
        r = dict_reader(oSheet, ("foo", "bar"), cell_typing=CellTyping.String)

        # verify
        self.assertEqual([
            {'bar': 'B1', 'foo': 'A1'},
            {'bar': 'B2', 'foo': 'A2'},
            {'bar': 'B3', 'foo': 'A3'}
        ], list(r))

    @patch("py4lo_io.get_cell_type")
    @patch("py4lo_io.get_used_range_address")
    def test_dict_reader_fieldnames_rest(self, gura, gct):
        # prepare
        oSheet = Mock()
        oSheet.getCellByPosition.side_effect = [
            Mock(String="A1"), Mock(), Mock(),
            Mock(String="A2"), Mock(String="B2"), Mock(String="C2"),
        ]
        gct.side_effect = ['TEXT', 'EMPTY', 'EMPTY', 'TEXT', 'TEXT', 'TEXT']
        oRangeAddress = Mock(StartColumn=0, StartRow=0, EndColumn=2, EndRow=1)
        gura.side_effect = [oRangeAddress]

        # play
        r = dict_reader(oSheet, ("foo", "bar"), restkey="RK", restval="RV",
                        cell_typing=CellTyping.Minimal)

        # verify
        self.assertEqual([
            {'bar': 'RV', 'foo': 'A1'}, {'RK': ['C2'], 'bar': 'B2', 'foo': 'A2'}
        ], list(r))


    def test_create_read_cell_other(self):
        with self.assertRaises(ValueError):
            create_read_cell(object())

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
