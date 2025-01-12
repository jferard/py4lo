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
import csv
import datetime as dt
import unittest
from typing import Any
from unittest import mock

from py4lo_helper import Target
from py4lo_io import (create_import_filter_options,
                      create_export_filter_options, Format, create_read_cell,
                      CellTyping, reader, dict_reader, find_number_format_style,
                      create_write_cell, writer, dict_writer, import_from_csv,
                      export_to_csv, Filter)
from py4lo_typing import UnoCell
from _mock_constants import NumberFormat


class Py4LOIOTestCase(unittest.TestCase):
    def test_create_read_cell_none(self):
        # prepare
        oCell = mock.Mock(String="foo")

        # play
        rc = create_read_cell(CellTyping.String)

        # verify
        self.assertEqual("foo", rc(oCell))

    @mock.patch("py4lo_io.get_cell_type")
    def test_create_read_cell_minimal_empty(self, gct):
        # prepare
        oCell = mock.Mock()
        gct.side_effect = ['EMPTY']

        # play
        rc = create_read_cell(CellTyping.Minimal)

        # verify
        self.assertIsNone(rc(oCell))

    @mock.patch("py4lo_io.get_cell_type")
    def test_create_read_cell_minimal_text(self, gct):
        # prepare
        oCell = mock.Mock(String="foo")
        gct.side_effect = ['TEXT']

        # play
        rc = create_read_cell(CellTyping.Minimal)

        # verify
        self.assertEqual("foo", rc(oCell))

    @mock.patch("py4lo_io.get_cell_type")
    def test_create_read_cell_minimal_value(self, gct):
        # prepare
        oCell = mock.Mock(Value=3.14)
        gct.side_effect = ['VALUE']

        # play
        rc = create_read_cell(CellTyping.Minimal)

        # verify
        self.assertEqual(3.14, rc(oCell))

    @mock.patch("py4lo_io.get_cell_type")
    def test_create_read_cell_minimal_other(self, gct):
        # prepare
        oCell = mock.Mock()
        gct.side_effect = ['OTHER']

        # play
        rc = create_read_cell(CellTyping.Minimal)

        # verify
        with self.assertRaises(ValueError):
            rc(oCell)

    @mock.patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_empty(self, gct):
        # prepare
        oCell = mock.Mock()
        oFormats = mock.Mock()
        gct.side_effect = ['EMPTY']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        self.assertIsNone(rc(oCell))

    @mock.patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_text(self, gct):
        # prepare
        oCell = mock.Mock(String="foo")
        oFormats = mock.Mock()
        gct.side_effect = ['TEXT']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        self.assertEqual("foo", rc(oCell))

    @mock.patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_date(self, gct):
        # prepare
        oCell = mock.Mock(Value=44000, NumberFormat=3)
        oFormats = mock.Mock()
        oFormats.getByKey.side_effect = [mock.Mock(Type=NumberFormat.DATE)]
        gct.side_effect = ['VALUE']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        self.assertEqual(dt.datetime(2020, 6, 18, 0, 0), rc(oCell))

    @mock.patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_logical(self, gct):
        # prepare
        oCell = mock.Mock(Value=1, NumberFormat=3)
        oFormats = mock.Mock()
        oFormats.getByKey.side_effect = [mock.Mock(Type=NumberFormat.LOGICAL)]
        gct.side_effect = ['VALUE']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        self.assertIs(True, rc(oCell))

    @mock.patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_number(self, gct):
        # prepare
        oCell = mock.Mock(Value=10, NumberFormat=3)
        oFormats = mock.Mock()
        oFormats.getByKey.side_effect = [mock.Mock(Type=NumberFormat.NUMBER)]
        gct.side_effect = ['VALUE']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        self.assertEqual(10, rc(oCell))

    @mock.patch("py4lo_io.get_cell_type")
    def test_create_read_cell_accurate_other(self, gct):
        # prepare
        oCell = mock.Mock()
        oFormats = mock.Mock()
        gct.side_effect = ['']

        # play
        rc = create_read_cell(CellTyping.Accurate, oFormats)

        # verify
        with self.assertRaises(ValueError):
            rc(oCell)

    def test_create_read_cell_accurate_wo_formats(self):
        with self.assertRaises(ValueError):
            create_read_cell(CellTyping.Accurate)

    def test_create_read_cell_other(self):
        with self.assertRaises(ValueError):
            create_read_cell(object())

    def test_reader_rc(self):
        # prepare
        oSheet = mock.Mock()
        oRange = mock.Mock(
            RangeAddress =self._range_address_mock(0, 0, 1, 2),
            Spreadsheet=oSheet
        )

        # play
        r = reader.from_read_cell(oRange, read_cell=lambda oCell: 15)

        # verify
        self.assertEqual([
            [15, 15],
            [15, 15],
            [15, 15]
        ], list(r))

    def test_reader(self):
        # prepare
        oSheet = mock.Mock()
        oRange = mock.Mock(
            RangeAddress=self._range_address_mock(0, 0, 1, 2),
            Spreadsheet=oSheet
        )
        oSheet.getCellByPosition.side_effect = [
            mock.Mock(String="A1"), mock.Mock(String="B1"),
            mock.Mock(String="A2"), mock.Mock(String="B2"),
            mock.Mock(String="A3"), mock.Mock(String="B3"),
        ]

        # play
        r = reader.from_typing(oRange, CellTyping.String)

        # verify
        self.assertEqual([
            ['A1', 'B1'], ['A2', 'B2'], ['A3', 'B3']
        ], list(r))

    @mock.patch("py4lo_io.get_cell_type")
    def test_reader_accurate(self, gct):
        # prepare
        gct.side_effect = ['TEXT', 'TEXT', 'TEXT', 'TEXT', 'TEXT', 'EMPTY']
        oSheet = mock.Mock()
        oRange = mock.Mock(
            RangeAddress=self._range_address_mock(0, 0, 1, 2),
            Spreadsheet =oSheet
        )
        oSheet.getCellByPosition.side_effect = [
            mock.Mock(String="A1"), mock.Mock(String="B1"),
            mock.Mock(String="A2"), mock.Mock(String="B2"),
            mock.Mock(String="A3"), mock.Mock(),
        ]
        oFormats = mock.Mock()

        # play
        r = reader.from_typing(oRange, CellTyping.Accurate, oFormats)

        # verify
        self.assertEqual([
            ['A1', 'B1'], ['A2', 'B2'], ['A3']
        ], list(r))

    def test_dict_reader(self):
        # prepare
        oSheet = mock.Mock()
        oRange = mock.Mock(
            RangeAddress =self._range_address_mock(0, 0, 1, 2),
            Spreadsheet=oSheet)
        oSheet.getCellByPosition.side_effect = [
            mock.Mock(String="A1"), mock.Mock(String="B1"),
            mock.Mock(String="A2"), mock.Mock(String="B2"),
            mock.Mock(String="A3"), mock.Mock(String="B3"),
        ]

        # play
        r = dict_reader.from_typing(oRange, cell_typing=CellTyping.String)

        # verify
        self.assertEqual([
            {'A1': 'A2', 'B1': 'B2'}, {'A1': 'A3', 'B1': 'B3'}
        ], list(r))

    def _range_address_mock(self, c1, r1, c2, r2):
        return mock.Mock(StartColumn=c1, StartRow=r1, EndColumn=c2, EndRow=r2)

    def test_dict_reader_fieldnames(self):
        # prepare
        oSheet = mock.Mock()
        oRange = mock.Mock(
            RangeAddress=self._range_address_mock(0, 0, 1, 2),
            Spreadsheet=oSheet,
        )
        oSheet.getCellByPosition.side_effect = [
            mock.Mock(String="A1"), mock.Mock(String="B1"),
            mock.Mock(String="A2"), mock.Mock(String="B2"),
            mock.Mock(String="A3"), mock.Mock(String="B3"),
        ]

        # play
        r = dict_reader.from_typing(oRange, ("foo", "bar"), cell_typing=CellTyping.String)

        # verify
        self.assertEqual([
            {'bar': 'B1', 'foo': 'A1'},
            {'bar': 'B2', 'foo': 'A2'},
            {'bar': 'B3', 'foo': 'A3'}
        ], list(r))

    @mock.patch("py4lo_io.get_cell_type")
    def test_dict_reader_fieldnames_rest(self, gct):
        # prepare
        oSheet = mock.Mock()
        oRange = mock.Mock(
            RangeAddress=self._range_address_mock(0, 0, 2, 1),
            Spreadsheet=oSheet
        )
        oSheet.getCellByPosition.side_effect = [
            mock.Mock(String="A1"), mock.Mock(), mock.Mock(),
            mock.Mock(String="A2"), mock.Mock(String="B2"),
            mock.Mock(String="C2"),
        ]
        gct.side_effect = ['TEXT', 'EMPTY', 'EMPTY', 'TEXT', 'TEXT', 'TEXT']

        # play
        r = dict_reader.from_typing(oRange, ("foo", "bar"), restkey="RK", restval="RV",
                        cell_typing=CellTyping.Accurate)

        # verify
        self.assertEqual([
            {'bar': 'RV', 'foo': 'A1'}, {'RK': ['C2'], 'bar': 'B2', 'foo': 'A2'}
        ], list(r))

    def test_find_number_format_style(self):
        # prepare
        oFormats = mock.Mock()
        oLocale = mock.Mock()
        oFormats.getStandardFormat.side_effect = [1]

        # play
        act_n = find_number_format_style(
            oFormats, NumberFormat.NUMBER, oLocale)

        # verify
        self.assertEqual(1, act_n)
        self.assertEqual([
            mock.call.getStandardFormat(NumberFormat.NUMBER, oLocale)],
            oFormats.mock_calls)

    def test_create_write_cell_string(self):
        # prepare
        oCell = mock.Mock()

        # play
        wc = create_write_cell(CellTyping.String)
        wc(oCell, 10)

        # verify
        self.assertEqual("10", oCell.String)

    def test_create_write_cell_minimal_none(self):
        # prepare
        oCell = mock.Mock()

        # play
        wc = create_write_cell(CellTyping.Minimal)
        wc(oCell, None)

        # verify
        self.assertEqual("", oCell.String)

    def test_create_write_cell_minimal_string(self):
        # prepare
        oCell = mock.Mock()

        # play
        wc = create_write_cell(CellTyping.Minimal)
        wc(oCell, "foo")

        # verify
        self.assertEqual("foo", oCell.String)

    def test_create_write_cell_minimal_date(self):
        # prepare
        oCell = mock.Mock()

        # play
        wc = create_write_cell(CellTyping.Minimal)
        wc(oCell, dt.date(2021, 1, 2))

        # verify
        self.assertEqual(44198.0, oCell.Value)

    def test_create_write_cell_minimal_bool(self):
        # prepare
        oCell = mock.Mock()

        # play
        wc = create_write_cell(CellTyping.Minimal)
        wc(oCell, True)

        # verify
        self.assertEqual(1, oCell.Value)

    def test_create_write_cell_minimal_number(self):
        # prepare
        oCell = mock.Mock()

        # play
        wc = create_write_cell(CellTyping.Minimal)
        wc(oCell, 10)

        # verify
        self.assertEqual(10.0, oCell.Value)

    def test_create_write_cell_accurate_none(self):
        # prepare
        oCell = mock.Mock()
        oFormats = mock.Mock()

        # play
        wc = create_write_cell(CellTyping.Accurate, oFormats)
        wc(oCell, None)

        # verify
        self.assertEqual("", oCell.String)

    def test_create_write_cell_accurate_string(self):
        # prepare
        oCell = mock.Mock()
        oFormats = mock.Mock()

        # play
        wc = create_write_cell(CellTyping.Accurate, oFormats)
        wc(oCell, "foo")

        # verify
        self.assertEqual("foo", oCell.String)

    def test_create_write_cell_accurate_number(self):
        # prepare
        oCell = mock.Mock()
        oFormats = mock.Mock()

        # play
        wc = create_write_cell(CellTyping.Accurate, oFormats)
        wc(oCell, 10)

        # verify
        self.assertEqual(10.0, oCell.Value)

    def test_create_write_cell_accurate_date(self):
        # prepare
        oCell = mock.Mock()
        oFormats = mock.Mock()
        oFormats.getStandardFormat.side_effect = [1, 2, 3]

        # play
        wc = create_write_cell(CellTyping.Accurate, oFormats)
        wc(oCell, dt.date(2020, 1, 1))

        # verify
        self.assertEqual(43831.0, oCell.Value)
        self.assertEqual(1, oCell.NumberFormat)

    def test_create_write_cell_accurate_datetime(self):
        # prepare
        oCell = mock.Mock()
        oFormats = mock.Mock()
        oFormats.getStandardFormat.side_effect = [1, 2, 3]

        # play
        wc = create_write_cell(CellTyping.Accurate, oFormats)
        wc(oCell, dt.datetime(2020, 1, 1, 4, 5, 6))

        # verify
        self.assertEqual(43831.17020833334, oCell.Value)
        self.assertEqual(2, oCell.NumberFormat)

    def test_create_write_cell_accurate_bool(self):
        # prepare
        oCell = mock.Mock()
        oFormats = mock.Mock()
        oFormats.getStandardFormat.side_effect = [1, 2, 3]

        # play
        wc = create_write_cell(CellTyping.Accurate, oFormats)
        wc(oCell, True)

        # verify
        self.assertEqual(1, oCell.Value)
        self.assertEqual(3, oCell.NumberFormat)

    def test_create_write_cell_accurate_no_format(self):
        with self.assertRaises(ValueError):
            create_write_cell(CellTyping.Accurate)

    def test_create_write_cell_other(self):
        with self.assertRaises(ValueError):
            create_write_cell(object())

    def test_writer_wc(self):
        # prepare
        oSheet = mock.Mock()
        cells = [mock.Mock() for _ in range(6)]
        oSheet.getCellByPosition.side_effect = cells

        # play
        def wc(oCell: UnoCell, value: Any):
            oCell.t = value

        w = writer(oSheet, write_cell=wc)
        w.writerows([
            ("a", "b", "oTextRange"),
            (1, 2, 3),
        ])

        #
        self.assertEqual([
            mock.call.getCellByPosition(0, 0),
            mock.call.getCellByPosition(1, 0),
            mock.call.getCellByPosition(2, 0),
            mock.call.getCellByPosition(0, 1),
            mock.call.getCellByPosition(1, 1),
            mock.call.getCellByPosition(2, 1)
        ], oSheet.mock_calls)
        self.assertEqual(['a', 'b', 'oTextRange', 1, 2, 3],
                         [c.t for c in cells])

    def test_writer_formats(self):
        # prepare
        oSheet = mock.Mock()
        cells = [mock.Mock() for _ in range(6)]
        oSheet.getCellByPosition.side_effect = cells
        oFormats = mock.Mock()

        # play
        w = writer.from_typing(oSheet, CellTyping.Accurate, oFormats=oFormats)
        w.writerows([
            ("a", "b", "oTextRange"),
            (1, 2, 3),
        ])

        #
        self.assertEqual([
            mock.call.getCellByPosition(0, 0),
            mock.call.getCellByPosition(1, 0),
            mock.call.getCellByPosition(2, 0),
            mock.call.getCellByPosition(0, 1),
            mock.call.getCellByPosition(1, 1),
            mock.call.getCellByPosition(2, 1)
        ], oSheet.mock_calls)
        self.assertEqual(['a', 'b', 'oTextRange'],
                         [c.String for c in cells[:3]])
        self.assertEqual([1, 2, 3], [c.Value for c in cells[3:]])

    def test_writer_no_formats(self):
        # prepare
        oSheet = mock.Mock()
        cells = [mock.Mock() for _ in range(6)]
        oSheet.getCellByPosition.side_effect = cells

        # play
        w = writer.from_typing(oSheet, CellTyping.Accurate)
        w.writerows([
            ("a", "b", "oTextRange"),
            (1, 2, 3),
        ])

        # verify
        self.assertEqual([
            mock.call.DrawPage.Forms.Parent.NumberFormats.getStandardFormat(
                2, mock.ANY),
            mock.call.DrawPage.Forms.Parent.NumberFormats.getStandardFormat(
                6, mock.ANY),
            mock.call.DrawPage.Forms.Parent.NumberFormats.getStandardFormat(
                1024, mock.ANY),
            mock.call.getCellByPosition(0, 0),
            mock.call.getCellByPosition(1, 0),
            mock.call.getCellByPosition(2, 0),
            mock.call.getCellByPosition(0, 1),
            mock.call.getCellByPosition(1, 1),
            mock.call.getCellByPosition(2, 1)
        ], oSheet.mock_calls)
        self.assertEqual(['a', 'b', 'oTextRange'],
                         [c.String for c in cells[:3]])
        self.assertEqual([1, 2, 3], [c.Value for c in cells[3:]])

    def test_dict_writer_wc(self):
        # prepare
        oSheet = mock.Mock()
        cells = [mock.Mock() for _ in range(9)]
        oSheet.getCellByPosition.side_effect = cells

        # play
        def wc(oCell: UnoCell, value: Any):
            oCell.t = value

        w = dict_writer.from_write_cell(oSheet, ['a', 'b', 'oTextRange'], write_cell=wc)

        w.writeheader()
        w.writerows([
            {"a": 1, "b": 2, "oTextRange": 3},
            {"a": 4, "b": 5, "oTextRange": 6},
        ])

        #
        self.assertEqual([
            mock.call.getCellByPosition(0, 0),
            mock.call.getCellByPosition(1, 0),
            mock.call.getCellByPosition(2, 0),
            mock.call.getCellByPosition(0, 1),
            mock.call.getCellByPosition(1, 1),
            mock.call.getCellByPosition(2, 1),
            mock.call.getCellByPosition(0, 2),
            mock.call.getCellByPosition(1, 2),
            mock.call.getCellByPosition(2, 2)
        ], oSheet.mock_calls)
        self.assertEqual(['a', 'b', 'oTextRange', 1, 2, 3, 4, 5, 6],
                         [c.t for c in cells])

    def test_dict_writer_wc_raise(self):
        # prepare
        oSheet = mock.Mock()
        cells = [mock.Mock() for _ in range(3)]
        oSheet.getCellByPosition.side_effect = cells

        # play
        def wc(oCell: UnoCell, value: Any):
            oCell.t = value

        w = dict_writer.from_write_cell(oSheet, ['a', 'b', 'oTextRange'], write_cell=wc)

        w.writeheader()
        with self.assertRaises(ValueError):
            w.writerow({"a": 1, "b": 2, "d": 3})

        #
        self.assertEqual([
            mock.call.getCellByPosition(0, 0),
            mock.call.getCellByPosition(1, 0),
            mock.call.getCellByPosition(2, 0),
        ], oSheet.mock_calls)
        self.assertEqual(['a', 'b', 'oTextRange'], [c.t for c in cells])

    def test_dict_writer_wc_restval(self):
        # prepare
        oSheet = mock.Mock()
        cells = [mock.Mock() for _ in range(6)]
        oSheet.getCellByPosition.side_effect = cells

        # play
        def wc(oCell: UnoCell, value: Any):
            oCell.t = value

        w = dict_writer.from_write_cell(oSheet, ['a', 'b', 'oTextRange'], restval="foo",
                        write_cell=wc)

        w.writeheader()
        w.writerow({"a": 1, "b": 2})

        #
        self.assertEqual([
            mock.call.getCellByPosition(0, 0),
            mock.call.getCellByPosition(1, 0),
            mock.call.getCellByPosition(2, 0),
            mock.call.getCellByPosition(0, 1),
            mock.call.getCellByPosition(1, 1),
            mock.call.getCellByPosition(2, 1),
        ], oSheet.mock_calls)
        self.assertEqual(['a', 'b', 'oTextRange', 1, 2, 'foo'],
                         [c.t for c in cells])


class IOCSVTestCase(unittest.TestCase):
    @mock.patch("py4lo_io.uno_path_to_url")
    @mock.patch("py4lo_io.get_provider")
    @mock.patch("py4lo_helper.make_pv")
    def test_import_from_csv(self, mkpv, get_provider, ptu):
        mkpv.side_effect = lambda a, b: (a, b)

        # prepare
        oSheets = mock.Mock()
        oDoc = mock.Mock(Sheets=oSheets)
        p = mock.Mock()
        ptu.side_effect = ["url"]
        oOtherSheets = mock.Mock(ElementNames=["a", "b", "oTextRange"])
        oOtherDoc = mock.Mock(Sheets=oOtherSheets)
        oDesktop = mock.Mock()
        oDesktop.loadComponentFromURL.side_effect = [oOtherDoc]
        provider = mock.Mock(desktop=oDesktop)
        get_provider.side_effect = [provider]

        # play
        import_from_csv(oDoc, "foo", 2, p, language_code="en_US")

        # verify
        self.assertEqual([mock.call.getByIndex(0)], oOtherSheets.mock_calls)
        self.assertEqual([mock.call.importSheet(oOtherDoc, 'a', 2)],
                         oSheets.mock_calls)
        self.assertEqual([
            mock.call.desktop.loadComponentFromURL(
                'url', Target.BLANK, 0, (
                    ("FilterName", Filter.CSV),
                    ("FilterOptions", "44,34,76,1,,1033,false,false"),
                    ("Hidden", True))
            )
        ], provider.mock_calls)

    def test_import_options_dialect(self):
        self.assertEqual('59,34,76,1,,1033,true,false',
                         create_import_filter_options(csv.unix_dialect,
                                                      delimiter=";",
                                                      language_code="en_US"))

    def test_import_options_two_args(self):
        with self.assertRaises(ValueError):
            create_import_filter_options(1, 2)

    def test_empty_import_options(self):
        self.assertEqual('44,34,76,1,,1033,false,false',
                         create_import_filter_options(
                             language_code="en_US"))

    def test_import_options(self):
        self.assertEqual('44,34,76,1,1/2,1033,false,false',
                         create_import_filter_options(language_code="en_US",
                                                      format_by_idx={
                                                          1: Format.TEXT}))

    @mock.patch("py4lo_io.uno_path_to_url")
    @mock.patch("py4lo_io.parent_doc")
    @mock.patch("py4lo_helper.make_pv")
    def test_export_to_csv(self, mkpv, pd, ptu):
        # prepare
        oSheet = mock.Mock()
        path = mock.Mock()
        ptu.side_effect = ['url']
        oCurSheet = mock.Mock()
        oDoc = mock.Mock(CurrentController=mock.Mock(ActiveSheet=oCurSheet))
        pd.side_effect = [oDoc]
        mkpv.side_effect = lambda a, b: (a, b)

        # play
        export_to_csv(oSheet, path, language_code="en_US")

        # verify
        self.assertEqual([], oSheet.mock_calls)
        self.assertEqual([
            mock.call.lockControllers(),
            mock.call.storeToURL(
                'url',
                (("FilterName", "Text - txt - csv (StarCalc)"),
                ("FilterOptions", "44,34,76,1,,1033,false,false,true"),
                ( "Overwrite", True))),
            mock.call.unlockControllers()
        ], oDoc.mock_calls)

    def test_empty_export_options(self):
        self.assertEqual('44,34,76,1,,1033,false,false,true',
                         create_export_filter_options(
                             language_code="en_US"))

    def test_export_options_dialect(self):
        self.assertEqual('44,39,76,1,,1033,true,true,true',
                         create_export_filter_options(csv.unix_dialect,
                                                      quotechar="'",
                                                      language_code="en_US"))

    def test_export_options_two_parameters(self):
        with self.assertRaises(ValueError):
            create_export_filter_options(1, 2)

if __name__ == '__main__':
    unittest.main()
