# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   FastODS is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   FastODS is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""
import unittest
from unittest import mock
from unittest.mock import *

from py4lo_helper import *
import py4lo_helper
from py4lo_helper import _ObjectProvider, _Inspector

# noinspection PyUnresolvedReferences
from com.sun.star.script.provider import ScriptFrameworkErrorException
# noinspection PyUnresolvedReferences
from com.sun.star.uno import RuntimeException as UnoRuntimeException


class TestHelper(unittest.TestCase):
    def setUp(self):
        self.xsc = Mock()
        self.doc = Mock()
        self.ctrl = Mock()
        self.frame = Mock()
        self.parent_win = Mock()
        self.msp = Mock()
        self.ctxt = Mock()
        self.sm = Mock()
        self.desktop = Mock()
        py4lo_helper.provider = _ObjectProvider(
            self.doc, self.ctrl, self.frame, self.parent_win, self.msp,
            self.ctxt, self.sm, self.desktop)
        py4lo_helper._inspect = _Inspector(py4lo_helper.provider)

    def test_init(self):
        xsc = Mock()
        init(xsc)
        self.assertIsNotNone(py4lo_helper.provider)
        self.assertIsNotNone(py4lo_helper._inspect)
        self.assertIsNotNone(py4lo_helper.xray)
        self.assertIsNotNone(py4lo_helper.mri)

    def test_get_script_provider_factory_twice(self):
        # prepare
        spf = Mock()
        self.sm.createInstanceWithContext.side_effect = [spf]

        # play
        actual_spf1 = py4lo_helper.provider.get_script_provider_factory()
        actual_spf2 = py4lo_helper.provider.get_script_provider_factory()

        # verify
        self.assertIs(spf, actual_spf1)
        self.assertIs(spf, actual_spf2)

    def test_get_script_provider(self):
        # prepare
        sp = Mock()
        spf = Mock()
        spf.createScriptProvider.side_effect = [sp]
        self.sm.createInstanceWithContext.side_effect = [spf]

        # play
        actual_sp = py4lo_helper.provider.get_script_provider()

        # verify
        self.assertIs(sp, actual_sp)

    def test_get_script_provider_twice(self):
        # prepare
        sp = Mock()
        spf = Mock()
        spf.createScriptProvider.side_effect = [sp]
        self.sm.createInstanceWithContext.side_effect = [spf]

        # play
        actual_sp1 = py4lo_helper.provider.get_script_provider()
        actual_sp2 = py4lo_helper.provider.get_script_provider()

        # verify
        self.assertIs(sp, actual_sp1)
        self.assertIs(sp, actual_sp2)

    def test_reflect(self):
        # prepare
        ref = Mock()
        self.sm.createInstance.side_effect = [ref]

        # play
        actual_ref = py4lo_helper.provider.reflect

        # verify
        self.assertEqual(
            [call.createInstance('com.sun.star.reflection.CoreReflection')],
            self.sm.mock_calls)
        self.assertEqual(ref, actual_ref)

    def test_reflect_twice(self):
        # prepare
        ref = Mock()
        self.sm.createInstance.side_effect = [ref]

        # play
        actual_ref1 = py4lo_helper.provider.reflect
        actual_ref2 = py4lo_helper.provider.reflect

        # verify
        self.assertEqual(
            [call.createInstance('com.sun.star.reflection.CoreReflection')],
            self.sm.mock_calls)
        self.assertEqual(ref, actual_ref1)
        self.assertEqual(ref, actual_ref2)

    def test_dispatcher(self):
        # prepare
        ref = Mock()
        self.sm.createInstance.side_effect = [ref]

        # play
        actual_ref = py4lo_helper.provider.dispatcher

        # verify
        self.assertEqual(
            [call.createInstance('com.sun.star.frame.DispatchHelper')],
            self.sm.mock_calls)
        self.assertEqual(ref, actual_ref)

    def test_dispatcher_twice(self):
        # prepare
        ref = Mock()
        self.sm.createInstance.side_effect = [ref]

        # play
        actual_ref1 = py4lo_helper.provider.dispatcher
        actual_ref2 = py4lo_helper.provider.dispatcher

        # verify
        self.assertEqual(
            [call.createInstance('com.sun.star.frame.DispatchHelper')],
            self.sm.mock_calls)
        self.assertEqual(ref, actual_ref1)
        self.assertEqual(ref, actual_ref2)

    def test_to_iter(self):
        # prepare
        index_access = Mock()
        index_access.getCount.side_effect = [3]
        index_access.getByIndex.side_effect = [1, 4, 9]

        # play
        ret = list(to_iter(index_access))

        # verify
        self.assertEqual([1, 4, 9], ret)

    def test_to_dict(self):
        # prepare
        name_access = Mock()
        name_access.getElementNames.side_effect = [("foo", "bar", "baz")]
        name_access.getByName.side_effect = [1, 4, 9]

        # play
        ret = to_dict(name_access)

        # verify
        self.assertEqual({'bar': 4, 'baz': 9, 'foo': 1}, ret)

    def test_parent_doc(self):
        # prepare
        oDoc = Mock()
        oSH = Mock(DrawPage=Mock(Forms=Mock(Parent=oDoc)))
        oRange = Mock(Spreadsheet=oSH)

        # play
        ret = parent_doc(oRange)

        # verify
        self.assertEqual(oDoc, ret)

    def test_get_cell_type(self):
        # prepare
        oCell = Mock(Type=Mock(value="foo"))

        # play
        ret = get_cell_type(oCell)

        # verify
        self.assertEqual("foo", ret)

    def test_get_cell_type_formula(self):
        # prepare
        oCell = Mock(Type=Mock(value="FORMULA"),
                     FormulaResultType=Mock(value="bar"))

        # play
        ret = get_cell_type(oCell)

        # verify
        self.assertEqual("bar", ret)

    def test_get_named_cells(self):
        # prepare
        oCells = Mock()
        oRange = Mock(ReferredCells=oCells)

        oRanges = Mock()
        oRanges.getByName.side_effect = [oRange]
        oDoc = Mock(NamedRanges=oRanges)

        # play
        ret = get_named_cells(oDoc, "foo")

        # verify
        self.assertEqual(oCells, ret)

    def test_get_named_cell(self):
        # prepare
        oCell = Mock()
        oCells = Mock()
        oCells.getCellByPosition.side_effect = [oCell]
        oRange = Mock(ReferredCells=oCells)

        oRanges = Mock()
        oRanges.getByName.side_effect = [oRange]
        oDoc = Mock(NamedRanges=oRanges)

        # play
        ret = get_named_cell(oDoc, "foo")

        # verify
        self.assertEqual(oCell, ret)

    def test_get_main_cell(self):
        # prepare
        oMCell = Mock()
        oCursor = Mock(RangeAddress=Mock(StartColumn=1, StartRow=5))
        oSheet = Mock()
        oSheet.createCursorByRange.side_effect = [oCursor]
        oSheet.getCellByPosition.side_effect = [oMCell]
        oCell = Mock(Spreadsheet=oSheet)

        # play
        ret = get_main_cell(oCell)

        # verify
        self.assertEqual(oMCell, ret)
        self.assertEqual([call.collapseToMergedArea()], oCursor.mock_calls)
        self.assertEqual([
            call.createCursorByRange(oCell),
            call.getCellByPosition(1, 5)
        ], oSheet.mock_calls)

    def test_open_in_calc(self):
        # prepare
        py4lo_helper.provider = Mock()

        # play
        open_in_calc("/fname", Target.SELF, FrameSearchFlag.AUTO, Hidden=True)
        # verify

        self.assertEqual([
            call.desktop.loadComponentFromURL(
                'file:///fname',
                Target.SELF, 0, (make_pv("Hidden", True),))],
            py4lo_helper.provider.mock_calls)

    def testXray(self):
        py4lo_helper._inspect.use_xray()
        self.msp.getScript.assert_called_once_with(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')

    def testXray2(self):
        py4lo_helper._inspect.xray(1)
        py4lo_helper._inspect.xray(2)
        self.msp.getScript.assert_called_once_with(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')
        self.msp.getScript.return_value.invoke.assert_has_calls(
            [call((1,), (), ()), call((2,), (), ())])

    def testPv(self):
        pv = make_pv("name", "value")
        self.assertTrue("uno.com.sun.star.beans.PropertyValue" in str(type(pv)))
        self.assertEqual("name", pv.Name)
        self.assertEqual("value", pv.Value)

    def testPvs(self):
        pvs = make_pvs({"name1": "value1", "name2": "value2"})
        pvs = sorted(pvs, key=lambda pv: pv.Name)
        self.assertEqual("name1", pvs[0].Name)
        self.assertEqual("value1", pvs[0].Value)
        self.assertEqual("name2", pvs[1].Name)
        self.assertEqual("value2", pvs[1].Value)

    def testUnoService(self):
        uno_service_ctxt("x")
        self.sm.createInstanceWithContext.assert_called_once_with("x",
                                                                  self.ctxt)

        uno_service_ctxt("y", [1, 2, 3])
        self.sm.createInstanceWithArgumentsAndContext.assert_called_once_with(
            "y", [1, 2, 3], self.ctxt)

        uno_service("z")
        self.sm.createInstance.assert_called_once_with("z")

    def test_doc_builder(self):
        d = doc_builder(NewDocumentUrl.Calc)
        d.build()
        py4lo_helper.provider.desktop.loadComponentFromURL.assert_called_once_with(
            NewDocumentUrl.Calc, Target.BLANK, 0, ())
        oDoc = py4lo_helper.provider.desktop.loadComponentFromURL.return_value
        oDoc.lockControllers.assert_called_once()
        oDoc.unlockControllers.assert_called_once()

    def test_doc_builder_sheet_names(self):
        oDoc = py4lo_helper.provider.desktop.loadComponentFromURL.return_value
        oDoc.Sheets.getCount.side_effect = [3, 3, 3, 3, 3]
        s1, s2, s3 = MagicMock(), MagicMock(), MagicMock()
        oDoc.Sheets.getByIndex.side_effect = [s1, s2, s3]

        d = doc_builder(NewDocumentUrl.Calc)
        d.sheet_names(list("abcdef"), expand_if_necessary=True)
        d.build()
        py4lo_helper.provider.desktop.loadComponentFromURL.assert_called_once_with(
            NewDocumentUrl.Calc, Target.BLANK, 0, ())
        oDoc.lockControllers.assert_called_once()
        oDoc.Sheets.getCount.assert_called()

        oDoc.Sheets.getByIndex.assert_has_calls([call(0), call(1), call(2)])
        s1.setName.assert_called_once_with("a")
        s2.setName.assert_called_once_with("b")
        s3.setName.assert_called_once_with("c")
        oDoc.Sheets.insertNewByName.assert_has_calls(
            [call("d", 3), call("e", 4), call("f", 5)])

        oDoc.unlockControllers.assert_called_once()

    def test_doc_builder_sheet_names2(self):
        oDoc = py4lo_helper.provider.desktop.loadComponentFromURL.return_value
        oDoc.Sheets.getCount.side_effect = [3, 3, 3, 3, 3, 2]
        s1, s2, s3 = MagicMock(), MagicMock(), MagicMock()
        oDoc.Sheets.getByIndex.side_effect = [s1, s2, s3, s3]

        d = doc_builder(NewDocumentUrl.Calc)
        d.sheet_names(list("ab"), trunc_if_necessary=True)
        d.build()
        py4lo_helper.provider.desktop.loadComponentFromURL.assert_called_once_with(
            NewDocumentUrl.Calc, Target.BLANK, 0, ())
        oDoc.lockControllers.assert_called_once()
        oDoc.Sheets.getCount.assert_called()

        oDoc.Sheets.getByIndex.assert_has_calls([call(0), call(1), call(2)])
        s1.setName.assert_called_once_with("a")
        s2.setName.assert_called_once_with("b")
        s3.getName.assert_called_once()
        oDoc.Sheets.removeByName.assert_called_once_with(
            s3.getName.return_value)

        oDoc.unlockControllers.assert_called_once()

    @patch("py4lo_helper.uno")
    def test_make_sort_field(self, uno):
        # prepare
        struct = Mock()
        uno.createUnoStruct.side_effect = [struct]

        # play
        make_sort_field(1, False)

        # verify
        self.assertEqual(1, struct.Field)
        self.assertFalse(struct.IsAscending)

    def test_used_range(self):
        # prepare
        oRangeAddress = Mock(
            StartColumn=2, StartRow=4, EndColumn=8, EndRow=16
        )
        oCursor = Mock(RangeAddress=oRangeAddress)
        oSheet = Mock()
        oSheet.createCursor.side_effect = [oCursor]

        # play
        oRange = get_used_range(oSheet)

        # verify
        self.assertEqual([
            call.gotoStartOfUsedArea(True), call.gotoEndOfUsedArea(True)
        ], oCursor.mock_calls)
        self.assertEqual([
            call.createCursor(), call.getCellRangeByPosition(2, 4, 8, 16)
        ], oSheet.mock_calls)

    def test_narrow_range_to_address(self):
        # prepare
        oSheet = Mock()
        oRangeAddress = Mock(
            StartColumn=1, StartRow=2, EndColumn=4, EndRow=8
        )
        # play
        oRange = narrow_range_to_address(oSheet, oRangeAddress)

        # verify
        self.assertEqual([call.getCellRangeByPosition(1, 2, 4, 8)],
                         oSheet.mock_calls)

    def test_get_range_size(self):
        # prepare
        oRange = Mock(RangeAddress=Mock(
            StartColumn=1, StartRow=2, EndColumn=4, EndRow=8
        ))
        # play
        w, h = get_range_size(oRange)

        # verify
        self.assertEqual(4, w)
        self.assertEqual(7, h)

    def test_data_array(self):
        data_array = [
            ("", "", "", "", "", "", "",),
            ("", "", "a", "", "", "b", "",),
            ("", "", "", "c", "d", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "e", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",)
        ]
        self.assertEqual(1, top_void_row_count(data_array))
        self.assertEqual(2, bottom_void_row_count(data_array))
        self.assertEqual(2, left_void_row_count(data_array))
        self.assertEqual(1, right_void_row_count(data_array))

    def test_data_array0(self):
        data_array = []
        self.assertEqual(0, top_void_row_count(data_array))
        self.assertEqual(0, bottom_void_row_count(data_array))
        self.assertEqual(0, left_void_row_count(data_array))
        self.assertEqual(0, right_void_row_count(data_array))

    def test_data_array1(self):
        data_array = [("",)]
        self.assertEqual(1, top_void_row_count(data_array))
        self.assertEqual(1, bottom_void_row_count(data_array))
        self.assertEqual(1, left_void_row_count(data_array))
        self.assertEqual(1, right_void_row_count(data_array))

    def test_data_array2h(self):
        data_array = [("", "")]
        self.assertEqual(1, top_void_row_count(data_array))
        self.assertEqual(1, bottom_void_row_count(data_array))
        self.assertEqual(2, left_void_row_count(data_array))
        self.assertEqual(2, right_void_row_count(data_array))

    def test_data_array2v(self):
        data_array = [("",), ("",)]
        self.assertEqual(2, top_void_row_count(data_array))
        self.assertEqual(2, bottom_void_row_count(data_array))
        self.assertEqual(1, left_void_row_count(data_array))
        self.assertEqual(1, right_void_row_count(data_array))

    def test_data_array2(self):
        data_array = [
            ("", "", "", "", "", "", "",),
            ("", "", "a", "", "", "b", "",),
            ("", "", "", "c", "d", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "e", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",)
        ]
        self.assertEqual(1, top_void_row_count(data_array))
        self.assertEqual(2, bottom_void_row_count(data_array))
        self.assertEqual(2, left_void_row_count(data_array))
        self.assertEqual(1, right_void_row_count(data_array))

    def test_data_array3(self):
        data_array = [
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
        ]
        self.assertEqual(5, top_void_row_count(data_array))
        self.assertEqual(5, bottom_void_row_count(data_array))
        self.assertEqual(7, left_void_row_count(data_array))
        self.assertEqual(7, right_void_row_count(data_array))

    @mock.patch("py4lo_helper.get_used_range_address")
    def test_narrow_range(self, gura):
        # prepare
        data_array = [
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
        ]
        oSheet = mock.Mock()
        oRange = mock.Mock(Spreadsheet=oSheet,
                           RangeAddress=mock.Mock(StartColumn=1, EndColumn=7,
                                                  StartRow=1, EndRow=5))
        gura.side_effect = [mock.Mock(StartColumn=0, EndColumn=60,
                                      StartRow=0, EndRow=50)]
        oNRange = mock.Mock(DataArray=data_array)
        oSheet.getCellRangeByPosition.side_effect = [oNRange]

        # play
        nr = narrow_range(oRange, True)

        # verify
        self.assertIsNone(nr)
        self.assertEqual([mock.call.getCellRangeByPosition(1, 1, 7, 5)],
                         oSheet.mock_calls)

    @mock.patch("py4lo_helper.get_used_range_address")
    def test_narrow_range_dont_narrow_data(self, gura):
        # prepare
        data_array = [
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",),
        ]
        oSheet = mock.Mock()
        oRange = mock.Mock(Spreadsheet=oSheet,
                           RangeAddress=mock.Mock(StartColumn=1, EndColumn=7,
                                                  StartRow=1, EndRow=5))
        gura.side_effect = [mock.Mock(StartColumn=0, EndColumn=60,
                                      StartRow=0, EndRow=50)]
        oNRange = mock.Mock(DataArray=data_array)
        oSheet.getCellRangeByPosition.side_effect = [oNRange]

        # play
        nr = narrow_range(oRange)

        # verify
        self.assertEqual(oNRange, nr)
        self.assertEqual([mock.call.getCellRangeByPosition(1, 1, 7, 5)],
                         oSheet.mock_calls)

    @mock.patch("py4lo_helper.get_used_range_address")
    def test_narrow_range2(self, gura):
        # prepare
        data_array = [
            ("", "", "", "", "", "", "",),
            ("", "x", "", "", "", "", "",),
            ("", "", "y", "", "", "", "",),
            ("", "", "", "z", "", "t", "",),
            ("", "", "", "", "", "", "",),
        ]
        oSheet = mock.Mock()
        oRange = mock.Mock(Spreadsheet=oSheet,
                           RangeAddress=mock.Mock(StartColumn=1, EndColumn=7,
                                                  StartRow=1, EndRow=5))
        gura.side_effect = [mock.Mock(StartColumn=0, EndColumn=60,
                                      StartRow=0, EndRow=50)]
        oNRange = mock.Mock(DataArray=data_array)
        oNRange2 = mock.Mock()
        oSheet.getCellRangeByPosition.side_effect = [oNRange, oNRange2]

        # play
        nr = narrow_range(oRange, True)

        # verify
        self.assertEqual(oNRange2, nr)
        self.assertEqual([mock.call.getCellRangeByPosition(1, 1, 7, 5),
                          mock.call.getCellRangeByPosition(2, 2, 6, 4)],
                         oSheet.mock_calls)

    @patch("py4lo_helper.parent_doc")
    def test_copy_range(self, pd):
        # prepare
        oRange = Mock()
        oController = Mock()
        oDoc = Mock(CurrentController=oController)
        pd.side_effect = [oDoc]
        oDisp = Mock()
        self.sm.createInstance.side_effect = [oDisp]
        oRanges = Mock()
        oDoc.createInstance.side_effect = [oRanges]

        # play
        copy_range(oRange)

        # verify
        self.assertEqual([
            call.select(oRange),
            call.select(oRanges),
        ], oController.mock_calls)
        self.assertEqual([
            call.executeDispatch(oController, '.uno:Copy', '', 0, [])
        ], oDisp.mock_calls)

    @patch("py4lo_helper.parent_doc")
    def test_paste_range(self, pd):
        # prepare
        oSheet = Mock()
        oDestAddress = Mock()
        oController = Mock()
        oDoc = Mock(CurrentController=oController)
        oRange = Mock()
        oRanges = Mock()
        oSheet.getCellByPosition.side_effect = [oRange]
        oDoc.createInstance.side_effect = [oRanges]
        pd.side_effect = [oDoc]
        oDisp = Mock()
        self.sm.createInstance.side_effect = [oDisp]

        # play
        paste_range(oSheet, oDestAddress)

        # verify
        self.assertEqual([
            call.select(oRange),
            call.select(oRanges),
        ], oController.mock_calls)
        self.assertEqual([
            call.executeDispatch(oController, '.uno:InsertContents', '', 0,
                                 make_pvs({
                                     "Flags": "SVDT", "FormulaCommand": 0,
                                     "SkipEmptyCells": False,
                                     "Transpose": False, "AsLink": False,
                                     "MoveMode": 4
                                 }))
        ], oDisp.mock_calls)

    @patch("py4lo_helper.parent_doc")
    def test_paste_range_formulas(self, pd):
        # prepare
        oSheet = Mock()
        oDestAddress = Mock()
        oController = Mock()
        oDoc = Mock(CurrentController=oController)
        oRange = Mock()
        oRanges = Mock()
        oSheet.getCellByPosition.side_effect = [oRange]
        oDoc.createInstance.side_effect = [oRanges]
        pd.side_effect = [oDoc]
        oDisp = Mock()
        self.sm.createInstance.side_effect = [oDisp]

        # play
        paste_range(oSheet, oDestAddress, True)

        # verify
        self.assertEqual([
            call.select(oRange),
            call.select(oRanges),
        ], oController.mock_calls)
        self.assertEqual([
            call.executeDispatch(oController, '.uno:Paste', '', 0, [])
        ], oDisp.mock_calls)

    @patch("py4lo_helper.get_used_range")
    def test_data_array(self, gura):
        # prepare
        oSheet = Mock()
        da = [("foo",), ("bar",), ("baz",)]
        oRange = Mock(DataArray=da)
        gura.side_effect = [oRange]
        # play
        act_da = data_array(oSheet)

        # verify
        self.assertEqual(da, act_da)

    @patch("py4lo_helper.get_used_range")
    def test_data_array_empty(self, gura):
        # prepare
        oSheet = Mock()
        da = []
        oRange = Mock(DataArray=da)
        gura.side_effect = [oRange]
        # play
        act_da = data_array(oSheet)

        # verify
        self.assertEqual(da, act_da)

    def test_set_validation_list(self):
        # prepare
        oVal = Mock()
        oCell = Mock(Validation=oVal)

        # play
        set_validation_list_by_cell(oCell, ["a", "b", "c"])

        # verify
        self.assertEqual(ValidationType.LIST, oVal.Type)
        self.assertEqual(2, oVal.ShowList)
        self.assertTrue(oVal.ShowErrorMessage)
        self.assertTrue(oVal.IgnoreBlankCells)
        self.assertEqual('"a";"b";"c"', oVal.Formula1)
        self.assertIsNotNone(oCell.String)

    def test_set_validation_list2(self):
        # prepare
        oVal = Mock()
        oCell = Mock(Validation=oVal)

        # play
        set_validation_list_by_cell(oCell, [1, 2, 3], "foo", False, True,
                                    False)

        # verify
        self.assertEqual(ValidationType.LIST, oVal.Type)
        self.assertEqual(1, oVal.ShowList)
        self.assertFalse(oVal.ShowErrorMessage)
        self.assertFalse(oVal.IgnoreBlankCells)
        self.assertEqual('1;2;3', oVal.Formula1)
        self.assertEqual("foo", oCell.String)

    def test_set_validation_list3(self):
        # prepare
        oVal = Mock()
        oCell = Mock(Validation=oVal)

        # play
        set_validation_list_by_cell(oCell, [1, "a", False])

        # verify
        self.assertEqual(ValidationType.LIST, oVal.Type)
        self.assertEqual(2, oVal.ShowList)
        self.assertTrue(oVal.ShowErrorMessage)
        self.assertTrue(oVal.IgnoreBlankCells)
        self.assertEqual('1;"a";False', oVal.Formula1)
        self.assertIsNotNone(oCell.String)

    @patch("py4lo_helper.update_pvs")
    def test_sort_range(self, up):
        # prepare
        oSD = Mock()
        oRange = Mock()
        oRange.createSortDescriptor.side_effect = [oSD]
        sort_fields = tuple([Mock(), Mock()])

        # play
        sort_range(oRange, sort_fields, True)

        # verify
        self.assertEqual([
            call.createSortDescriptor(),
            call.sort(oSD)
        ], oRange.mock_calls)
        self.assertEqual([
            call(oSD, {'ContainsHeader': True, 'SortFields': ANY})
        ], up.mock_calls)

    @patch("py4lo_helper.update_pvs")
    def test_sort_range_no_header(self, up):
        # prepare
        oSD = Mock()
        oRange = Mock()
        oRange.createSortDescriptor.side_effect = [oSD]
        sort_fields = tuple([Mock(), Mock()])

        # play
        sort_range(oRange, sort_fields, False)

        # verify
        self.assertEqual([
            call.createSortDescriptor(),
            call.sort(oSD)
        ], oRange.mock_calls)
        self.assertEqual([
            call(oSD, {'ContainsHeader': False, 'SortFields': ANY})
        ], up.mock_calls)

    def test_clear_conditional_format(self):
        # prepare
        oFormat = Mock()
        oRange = Mock(ConditionalFormat=oFormat)

        # play
        clear_conditional_format(oRange)

        # verify
        self.assertEqual([call.clear()], oFormat.mock_calls)

    @patch("py4lo_helper.get_formula_conditional_entry")
    def test_conditional_format_on_formulas(self, gfce):
        # prepare
        oFormat = Mock()
        oRange = Mock(ConditionalFormat=oFormat)
        oCellAddress = Mock()
        gfce.side_effect = ["e"]

        # play
        conditional_format_on_formulas(
            oRange, {"foo": "bar"}, oCellAddress)

        # verify
        self.assertEqual([call('foo', 'bar', oCellAddress)], gfce.mock_calls)
        self.assertEqual([call.addNew("e")], oFormat.mock_calls)

    @patch("py4lo_helper.make_pv")
    def test_get_formula_conditional_entry(self, mk_pv):
        # prepare
        oCellAddress = Mock()
        mk_pv.side_effect = [1, 2, 3, 4, 5]

        # play
        ret = get_formula_conditional_entry("A1=0", "foo", oCellAddress)

        # verify
        self.assertEqual((1, 2, 3, 4, 5), ret)
        self.assertEqual([
            call('Formula1', 'A1=0'),
            call('Formula2', ''),
            call('Operator', ConditionOperator.FORMULA),
            call('StyleName', 'foo'),
            call('SourcePosition', oCellAddress)
        ], mk_pv.mock_calls)

    def test_find_number_format_style(self):
        # prepare
        oFormats = Mock()
        oFormats.queryKey.side_effect = [9]
        oDoc = Mock(NumberFormats=oFormats)
        oLocale = Mock()

        # play
        ret = find_or_create_number_format_style(oDoc, "YY-MM-DD", oLocale)

        # verify
        self.assertEqual(9, ret)

    @patch("py4lo_helper.make_locale")
    def test_find_number_format_style_no_locale(self, ml):
        # prepare
        oFormats = Mock()
        oFormats.queryKey.side_effect = [9]
        oDoc = Mock(NumberFormats=oFormats)
        oLocale = Mock()
        ml.side_effect = [oLocale]

        # play
        ret = find_or_create_number_format_style(oDoc, "YY-MM-DD")

        # verify
        self.assertEqual(9, ret)

    def test_create_number_format_style(self):
        # prepare
        oFormats = Mock()
        oFormats.queryKey.side_effect = [-1]
        oFormats.addNew.side_effect = [10]
        oDoc = Mock(NumberFormats=oFormats)
        oLocale = Mock()

        # play
        ret = find_or_create_number_format_style(oDoc, "YY-MM-DD", oLocale)

        # verify
        self.assertEqual(10, ret)

    @patch("py4lo_helper.parent_doc")
    def test_create_filter(self, pd):
        # prepare
        oFrame = Mock()
        oController = Mock(Frame=oFrame)
        oDoc = Mock(CurrentController=oController)
        oRange = Mock()
        pd.side_effect = [oDoc]

        # play
        create_filter(oRange)

        # verify
        self.assertEqual([
            call.select(oRange),
            call.select(ANY)
        ], oController.mock_calls)
        self.assertEqual([
            call.createInstance('com.sun.star.frame.DispatchHelper'),
            call.createInstance().executeDispatch(
                oFrame, '.uno:DataFilterAutoFilter', '', 0, [])
        ], self.sm.mock_calls)

    def test_row_as_header(self):
        # prepare
        oRow = Mock()

        # play
        row_as_header(oRow)

        # verify
        self.assertEqual(150, oRow.CharWeight)
        self.assertEqual(150, oRow.CharWeightAsian)
        self.assertEqual(150, oRow.CharWeightComplex)
        self.assertTrue(oRow.IsTextWrapped)
        self.assertTrue(oRow.OptimalHeight)

    def test_column_optimal_width_small(self):
        # prepare
        oColumn = Mock(Width=1 * 1000)

        # play
        column_optimal_width(oColumn)

        # verify
        self.assertFalse(oColumn.OptimalWidth)
        self.assertEqual(2 * 1000, oColumn.Width)

    def test_column_optimal_width_large(self):
        # prepare
        oColumn = Mock(Width=20 * 1000)

        # play
        column_optimal_width(oColumn)

        # verify
        self.assertFalse(oColumn.OptimalWidth)
        self.assertTrue(oColumn.IsTextWrapped)
        self.assertEqual(10 * 1000, oColumn.Width)

    def test_column_optimal_width_medium(self):
        # prepare
        oColumn = Mock(Width=5 * 1000)

        # play
        column_optimal_width(oColumn)

        # verify
        self.assertTrue(oColumn.OptimalWidth)

    @patch("py4lo_helper.get_used_range_address")
    def test_set_print_area(self, gura):
        # prepare
        sheet_ra = Mock()
        gura.side_effect = [sheet_ra]
        oSheet = Mock()
        title_ra = Mock()
        oRow = Mock(RangeAddress=title_ra)

        # play
        set_print_area(oSheet, oRow)

        # verify
        self.assertEqual([
            call.setPrintAreas([sheet_ra]),
            call.setPrintTitleRows(True),
            call.setTitleRows(title_ra)
        ], oSheet.mock_calls)

    @patch("py4lo_helper.get_used_range_address")
    def test_set_print_area_no_title(self, gura):
        # prepare
        sheet_ra = Mock()
        gura.side_effect = [sheet_ra]
        oSheet = Mock()

        # play
        set_print_area(oSheet)

        # verify
        self.assertEqual([
            call.setPrintAreas([sheet_ra])
        ], oSheet.mock_calls)

    @patch("py4lo_helper.parent_doc")
    def test_get_page_style(self, pd):
        # prepare
        oSheet = Mock(PageStyle="foo")
        oDoc = Mock()
        pd.side_effect = [oDoc]

        # play
        actual_ps = get_page_style(oSheet)

        # verify
        self.assertEqual([
            call.StyleFamilies.getByName('PageStyles'),
            call.StyleFamilies.getByName().getByName('foo')
        ], oDoc.mock_calls)

    @patch("py4lo_helper.make_struct")
    @patch("py4lo_helper.get_used_range")
    @patch("py4lo_helper.get_page_style")
    def test_set_paper_1(self, gps, gur, ms):
        self._set_paper(gps, gur, ms, 10 * 1000, 40 * 1000, False, 1, 0, 21000,
                        29700)

    @patch("py4lo_helper.make_struct")
    @patch("py4lo_helper.get_used_range")
    @patch("py4lo_helper.get_page_style")
    def test_set_paper_2(self, gps, gur, ms):
        self._set_paper(gps, gur, ms, 40 * 1000, 400 * 1000, False, 1, 0,
                        29700, 42000)

    @patch("py4lo_helper.make_struct")
    @patch("py4lo_helper.get_used_range")
    @patch("py4lo_helper.get_page_style")
    def test_set_paper_3(self, gps, gur, ms):
        self._set_paper(gps, gur, ms, 40 * 1000, 10 * 1000, True, 0, 1, 29700,
                        21000)

    @patch("py4lo_helper.make_struct")
    @patch("py4lo_helper.get_used_range")
    @patch("py4lo_helper.get_page_style")
    def test_set_paper_4(self, gps, gur, ms):
        self._set_paper(gps, gur, ms, 400 * 1000, 40 * 1000, True, 0, 1, 42000,
                        29700)

    def _set_paper(self, gps, gur, ms, w, h, is_landscape, sx, sy, pw, ph):
        # prepare
        oSheet = Mock()
        oPageStyle = Mock()
        gps.side_effect = [oPageStyle]
        oRange = Mock(Size=Mock(Width=w, Height=h))
        gur.side_effect = [oRange]
        # play
        set_paper(oSheet)
        # verify
        self.assertEqual(is_landscape, oPageStyle.IsLandscape)
        self.assertEqual(sx, oPageStyle.ScaleToPagesX)
        self.assertEqual(sy, oPageStyle.ScaleToPagesY)
        self.assertEqual([call('com.sun.star.awt.Size', Width=pw, Height=ph)],
                         ms.mock_calls)

    @patch("py4lo_helper.parent_doc")
    def test_add_link(self, pd):
        # prepare
        oText = Mock(End=1)
        oCur = Mock()
        oText.createTextCursorByRange.side_effect = [oCur]
        oCell = Mock(Text=oText)
        oDoc = Mock()
        oURL = Mock()
        oDoc.createInstance.side_effect = [oURL]
        pd.side_effect = [oDoc]

        # play
        add_link(oCell, "a very long text", "https://url.org", -1)

        # verify
        self.assertEqual([
            call.Text.createTextCursorByRange(1),
            call.insertTextContent(oCur, oURL, False)
        ], oCell.mock_calls)
        self.assertEqual('a very long text', oURL.Representation)
        self.assertEqual("https://url.org", oURL.URL)

    @patch("py4lo_helper.parent_doc")
    def test_add_link_wrapped(self, pd):
        # prepare
        oText = Mock(End=1)
        oCur = Mock()
        oText.createTextCursorByRange.side_effect = [oCur]
        oCell = Mock(Text=oText)
        oDoc = Mock()
        oURL1 = Mock()
        oURL2 = Mock()
        oDoc.createInstance.side_effect = [oURL1, oURL2]
        pd.side_effect = [oDoc]

        # play
        add_link(oCell, "a very long text", "https://url.org", 10)

        # verify
        self.assertEqual([
            call.Text.createTextCursorByRange(1),
            call.insertTextContent(oCur, oURL1, False),
            call.insertTextContent(oCur, oURL2, False)
        ], oCell.mock_calls)
        self.assertEqual('a very long', oURL1.Representation)
        self.assertEqual("https://url.org", oURL1.URL)
        self.assertEqual('text', oURL2.Representation)
        self.assertEqual("https://url.org", oURL2.URL)

    def test_wrap_text(self):
        self.assertEqual(['abcd', 'efgh', 'ij'],
                         py4lo_helper._wrap_text("abcd efgh ij", 6))
        self.assertEqual(['abcd', 'efgh ij'],
                         py4lo_helper._wrap_text("abcd efgh ij", 7))
        self.assertEqual(['abcd', 'efgh ij'],
                         py4lo_helper._wrap_text("abcd efgh ij", 8))
        self.assertEqual(['abcd efgh', 'ij'],
                         py4lo_helper._wrap_text("abcd efgh ij", 9))
        self.assertEqual(['abcd efgh', 'ij'],
                         py4lo_helper._wrap_text("abcd efgh ij", 10))
        self.assertEqual(['abcd efgh ij'],
                         py4lo_helper._wrap_text("abcd efgh ij", 11))
        self.assertEqual(['a very', 'long', 'text'],
                         py4lo_helper._wrap_text("a very long text", 8))
        self.assertEqual(['a very', 'long text'],
                         py4lo_helper._wrap_text("a very long text", 9))

    def test_read_options(self):
        # prepare
        oRange = Mock(DataArray=[
            ("foo", 1, "", ""),
            ("bar", 1, 2, 3),
            ("baz", "a", "", "b"),
            ("", "x", "", ""),
        ])
        oSheet = Mock()
        oSheet.getCellRangeByPosition.side_effect = [oRange]
        aAddress = Mock()
        aAddress.StartColumn = 0
        aAddress.EndColumn = 1
        aAddress.StartRow = 0
        aAddress.EndRow = 3

        def f(k, v): return k, v

        # play
        act_options = read_options(oSheet, aAddress, f)

        # verify
        self.assertEqual({'foo': 1, 'bar': (1, 2, 3), 'baz': ('a', '', 'b')},
                         act_options)

    def test_read_empty_options(self):
        # prepare
        oSheet = Mock()
        aAddress = Mock()
        aAddress.StartColumn = 0
        aAddress.EndColumn = 0
        aAddress.StartRow = 0
        aAddress.EndRow = 0

        def f(k, v): return k, v

        # play
        act_options = read_options(oSheet, aAddress, f)

        # verify
        self.assertEqual({}, act_options)

    def test_rtrim_row(self):
        self.assertEqual("", rtrim_row(tuple()))
        self.assertEqual(None, rtrim_row(tuple(), None))
        self.assertEqual("", rtrim_row(("", "", "", "")))
        self.assertEqual(None, rtrim_row(("", "", "", ""), None))
        self.assertEqual("foo", rtrim_row(("foo", "", "", "")))
        self.assertEqual(0.0, rtrim_row((0.0, "", "", "")))
        self.assertEqual((0.0, "", 1.0), rtrim_row((0.0, "", 1.0, "")))
        self.assertEqual(("foo", "", "", "bar"),
                         rtrim_row(("foo", "", "", "bar")))

    @patch("py4lo_helper.read_options")
    @patch("py4lo_helper.get_used_range_address")
    def test_read_options_from_sheet_name(self, gura, ro):
        # prepare
        oSheet = Mock()
        oSheets = Mock()
        oSheets.getByName.side_effect = [oSheet]
        oDoc = Mock(Sheets=oSheets)
        oAddress = Mock()
        gura.side_effect = [oAddress]

        def f(k, v): return k, v

        # play
        read_options_from_sheet_name(oDoc, "foo", f)

        # verify
        self.assertEqual([call(oSheet, oAddress, f)], ro.mock_calls)

    def test_copy_row_at_index(self):
        # prepare
        oRange = Mock()
        oSheet = Mock()
        oSheet.getCellRangeByPosition.side_effect = [oRange]
        row = ("foo", "bar", "baz")

        # play
        copy_row_at_index(oSheet, row, 3)

        # verifiy
        self.assertEqual([call.getCellRangeByPosition(0, 3, 2, 3)],
                         oSheet.mock_calls)
        self.assertEqual(row, oRange.DataArray)


class InspectorTestCase(unittest.TestCase):
    def setUp(self):
        self.oSP = Mock()
        self.provider = Mock(script_provider=self.oSP)
        self.inspector = _Inspector(self.provider)

    def test_use_xray_ok(self):
        # prepare
        # play
        self.inspector.use_xray(False)

        # verify
        self.assertEqual([call.getScript(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')],
            self.oSP.mock_calls)
        self.assertFalse(self.inspector._ignore_xray)

    def test_use_xray_raise(self):
        # prepare

        self.oSP.getScript.side_effect = ScriptFrameworkErrorException

        # play
        with self.assertRaises(UnoRuntimeException):
            self.inspector.use_xray(True)

        # verify
        self.assertEqual([call.getScript(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')],
            self.oSP.mock_calls)
        self.assertTrue(self.inspector._ignore_xray)

    def test_use_xray_ignore(self):
        # prepare
        self.oSP.getScript.side_effect = ScriptFrameworkErrorException

        # play
        self.inspector.use_xray(False)

        # verify
        self.assertEqual([call.getScript(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')],
            self.oSP.mock_calls)
        self.assertTrue(self.inspector._ignore_xray)

    def test_xray(self):
        # prepare
        obj = Mock()

        # play
        self.inspector.xray(obj)

        # verify
        self.assertEqual([
            call.getScript(
                'vnd.sun.star.script:XrayTool._Main.Xray?'
                'language=Basic&location=application'),
            call.getScript().invoke((obj,), tuple(), tuple())
        ], self.oSP.mock_calls)
        self.assertFalse(self.inspector._ignore_xray)

    def test_xray_fail(self):
        # prepare
        self.oSP.getScript.side_effect = ScriptFrameworkErrorException
        obj = Mock()

        # play
        self.inspector.xray(obj)

        # verify
        self.assertEqual([call.getScript(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')],
            self.oSP.mock_calls)
        self.assertTrue(self.inspector._ignore_xray)

    def test_xray_twice(self):
        # prepare
        self.oSP.getScript.side_effect = ScriptFrameworkErrorException
        obj = Mock()

        # play
        self.inspector.use_xray(False)
        self.inspector.xray(obj)

        # verify
        self.assertEqual([call.getScript(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')],
            self.oSP.mock_calls)
        self.assertTrue(self.inspector._ignore_xray)

    @patch("py4lo_helper.uno_service")
    def test_mri(self, us):
        # prepare
        obj = Mock()
        oMRI = Mock()
        us.side_effect = [oMRI]

        # play
        self.inspector.mri(obj)

        # verify
        self.assertEqual([
            call.inspect(obj)
        ], oMRI.mock_calls)
        self.assertFalse(self.inspector._ignore_mri)

    @patch("py4lo_helper.uno_service")
    def test_mri_twice(self, us):
        # prepare
        obj = Mock()
        oMRI = Mock()
        us.side_effect = [oMRI]

        # play
        self.inspector.mri(obj)
        self.inspector.mri(obj)

        # verify
        self.assertEqual([
            call.inspect(obj),
            call.inspect(obj)
        ], oMRI.mock_calls)
        self.assertFalse(self.inspector._ignore_mri)

    @patch("py4lo_helper.uno_service")
    def test_mri_fail(self, us):
        # prepare
        obj = Mock()
        us.side_effect = ScriptFrameworkErrorException

        # play
        self.inspector.mri(obj)

        # verify
        self.assertTrue(self.inspector._ignore_mri)

    @patch("py4lo_helper.uno_service")
    def test_mri_fail_twice(self, us):
        # prepare
        obj = Mock()
        us.side_effect = ScriptFrameworkErrorException

        # play
        self.inspector.mri(obj)
        self.inspector.mri(obj)

        # verify
        self.assertTrue(self.inspector._ignore_mri)

    @patch("py4lo_helper.uno_service")
    def test_mri_fail_err(self, us):
        # prepare
        obj = Mock()
        us.side_effect = ScriptFrameworkErrorException

        # play
        with self.assertRaises(UnoRuntimeException):
            self.inspector.mri(obj, True)

        # verify
        self.assertTrue(self.inspector._ignore_mri)


if __name__ == '__main__':
    unittest.main()
