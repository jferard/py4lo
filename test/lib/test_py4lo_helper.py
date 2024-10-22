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
# mypy: disable-error-code="import-not-found"
import unittest
from unittest import mock

import py4lo_helper
from py4lo_helper import (
    BorderLineStyle, ValidationType, ConditionOperator,
    FrameSearchFlag, ScriptFrameworkErrorException, UnoRuntimeException,
    to_iter, to_enumerate, to_dict, parent_doc, get_cell_type, get_named_cell,
    get_named_cells, get_main_cell,
    init, _ObjectProvider, _Inspector, PropertyState, create_uno_struct,
    make_pv, make_full_pv, make_pvs, update_pvs, make_locale, make_border,
    make_sort_field, get_last_used_row, get_used_range_address, get_used_range,
    narrow_range_to_address, get_range_size, copy_range, paste_range,
    narrow_range, top_void_row_count, bottom_void_row_count,
    left_void_row_count, right_void_row_count, data_array,
    ListValidationBuilder, set_validation_list_by_cell, sort_range,
    quote_element, clear_conditional_format, conditional_format_on_formulas,
    get_formula_conditional_entry, find_or_create_number_format_style,
    create_filter, row_as_header, column_optimal_width, set_print_area,
    get_page_style, set_paper, add_link, _wrap_text, open_in_calc, Target,
    new_doc, NewDocumentUrl, doc_builder, create_uno_service_ctxt,
    create_uno_service, read_options, rtrim_row, read_options_from_sheet_name,
    copy_row_at_index, FontSlant)
from py4lo_typing import UnoTextRange


# noinspection PyUnresolvedReferences
# from com.sun.star.script.provider import ScriptFrameworkErrorException
# noinspection PyUnresolvedReferences
# from com.sun.star.uno import RuntimeException as UnoRuntimeException


class HelperBaseTestCase(unittest.TestCase):
    def setUp(self):
        self.xsc = mock.Mock()
        self.doc = mock.Mock()
        self.ctrl = mock.Mock()
        self.frame = mock.Mock()
        self.parent_win = mock.Mock()
        self.msp = mock.Mock()
        self.ctxt = mock.Mock()
        self.sm = mock.Mock()
        self.desktop = mock.Mock()
        self.provider = _ObjectProvider(self.doc, self.ctrl, self.frame,
                                        self.parent_win, self.msp, self.ctxt,
                                        self.sm, self.desktop)
        py4lo_helper.provider = self.provider
        py4lo_helper._inspect = _Inspector(py4lo_helper.provider)

    def test_init(self):
        xsc = mock.Mock()
        init(xsc)
        self.assertIsNotNone(py4lo_helper.provider)
        self.assertIsNotNone(py4lo_helper._inspect)
        self.assertIsNotNone(py4lo_helper.xray)
        self.assertIsNotNone(py4lo_helper.mri)

    def test_get_script_provider_factory_twice(self):
        # prepare
        spf = mock.Mock()
        self.sm.createInstanceWithContext.side_effect = [spf]

        # play
        actual_spf1 = self.provider.get_script_provider_factory()
        actual_spf2 = self.provider.get_script_provider_factory()

        # verify
        self.assertIs(spf, actual_spf1)
        self.assertIs(spf, actual_spf2)

    def test_get_script_provider(self):
        # prepare
        sp = mock.Mock()
        spf = mock.Mock()
        spf.createScriptProvider.side_effect = [sp]
        self.sm.createInstanceWithContext.side_effect = [spf]

        # play
        actual_sp = py4lo_helper.provider.get_script_provider()

        # verify
        self.assertIs(sp, actual_sp)

    def test_get_script_provider_twice(self):
        # prepare
        sp = mock.Mock()
        spf = mock.Mock()
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
        ref = mock.Mock()
        self.sm.createInstance.side_effect = [ref]

        # play
        actual_ref = py4lo_helper.provider.reflect

        # verify
        self.assertEqual(
            [
                mock.call.createInstance(
                    'com.sun.star.reflection.CoreReflection')
            ],
            self.sm.mock_calls)
        self.assertEqual(ref, actual_ref)

    def test_reflect_twice(self):
        # prepare
        ref = mock.Mock()
        self.sm.createInstance.side_effect = [ref]

        # play
        actual_ref1 = py4lo_helper.provider.reflect
        actual_ref2 = py4lo_helper.provider.reflect

        # verify
        self.assertEqual(
            [
                mock.call.createInstance(
                    'com.sun.star.reflection.CoreReflection')
            ],
            self.sm.mock_calls)
        self.assertEqual(ref, actual_ref1)
        self.assertEqual(ref, actual_ref2)

    def test_dispatcher(self):
        # prepare
        ref = mock.Mock()
        self.sm.createInstance.side_effect = [ref]

        # play
        actual_ref = py4lo_helper.provider.dispatcher

        # verify
        self.assertEqual(
            [
                mock.call.createInstance('com.sun.star.frame.DispatchHelper')
            ],
            self.sm.mock_calls)
        self.assertEqual(ref, actual_ref)

    def test_dispatcher_twice(self):
        # prepare
        ref = mock.Mock()
        self.sm.createInstance.side_effect = [ref]

        # play
        actual_ref1 = py4lo_helper.provider.dispatcher
        actual_ref2 = py4lo_helper.provider.dispatcher

        # verify
        self.assertEqual(
            [
                mock.call.createInstance('com.sun.star.frame.DispatchHelper')
            ],
            self.sm.mock_calls)
        self.assertEqual(ref, actual_ref1)
        self.assertEqual(ref, actual_ref2)

    def test_to_iter_count(self):
        # prepare
        oIndexAccess = mock.Mock()
        oIndexAccess.Types = [mock.Mock(typeName="com.sun.star.container.XIndexAccess")]
        oIndexAccess.Count = 3
        oIndexAccess.getByIndex.side_effect = [1, 4, 9]

        # play
        ret = list(to_iter(oIndexAccess))

        # verify
        self.assertEqual([1, 4, 9], ret)

    def test_to_enumerate_count(self):
        # prepare
        oIndexAccess = mock.Mock()
        oIndexAccess.Types = [mock.Mock(typeName="com.sun.star.container.XIndexAccess")]
        oIndexAccess.Count = 3
        oIndexAccess.getByIndex.side_effect = [1, 4, 9]

        # play
        ret = list(to_enumerate(oIndexAccess))

        # verify
        self.assertEqual([(0, 1), (1, 4), (2, 9)], ret)

    def test_to_iter_enum(self):
        # prepare
        oEnum = mock.Mock()
        oEnumAccess = mock.Mock(spec=["createEnumeration"])
        oEnumAccess.Types = [mock.Mock(typeName="com.sun.star.container.XEnumerationAccess")]
        oEnumAccess.createEnumeration.side_effect = [oEnum]
        oEnum.hasMoreElements.side_effect = [True, True, True, False]
        oEnum.nextElement.side_effect = [1, 4, 9]

        # play
        ret = list(to_iter(oEnumAccess))

        # verify
        self.assertEqual([1, 4, 9], ret)

    def test_to_enumerate_enum(self):
        # prepare
        oEnum = mock.Mock()
        oEnumAccess = mock.Mock(spec=["createEnumeration"])
        oEnumAccess.Types = [mock.Mock(typeName="com.sun.star.container.XEnumerationAccess")]
        oEnumAccess.createEnumeration.side_effect = [oEnum]
        oEnum.hasMoreElements.side_effect = [True, True, True, False]
        oEnum.nextElement.side_effect = [1, 4, 9]

        # play
        ret = list(to_enumerate(oEnumAccess))

        # verify
        self.assertEqual([(0, 1), (1, 4), (2, 9)], ret)

    def test_to_dict(self):
        # prepare
        name_access = mock.Mock()
        name_access.ElementNames = ["foo", "bar", "baz"]
        name_access.getByName.side_effect = [1, 4, 9]

        # play
        ret = to_dict(name_access)

        # verify
        self.assertEqual({'bar': 4, 'baz': 9, 'foo': 1}, ret)

    def test_parent_doc(self):
        # prepare
        oDoc = mock.Mock()
        oSH = mock.Mock(DrawPage=mock.Mock(Forms=mock.Mock(Parent=oDoc)))
        oRange = mock.Mock(Spreadsheet=oSH)

        # play
        ret = parent_doc(oRange)

        # verify
        self.assertEqual(oDoc, ret)

    def test_get_cell_type(self):
        # prepare
        oCell = mock.Mock(Type=mock.Mock(value="foo"))

        # play
        ret = get_cell_type(oCell)

        # verify
        self.assertEqual("foo", ret)

    def test_get_cell_type_formula(self):
        # prepare
        oCell = mock.Mock(Type=mock.Mock(value="FORMULA"),
                          FormulaResultType=mock.Mock(value="bar"))

        # play
        ret = get_cell_type(oCell)

        # verify
        self.assertEqual("bar", ret)

    def test_get_named_cells(self):
        # prepare
        oCells = mock.Mock()
        oRange = mock.Mock(ReferredCells=oCells)

        oRanges = mock.Mock()
        oRanges.getByName.side_effect = [oRange]
        oDoc = mock.Mock(NamedRanges=oRanges)

        # play
        ret = get_named_cells(oDoc, "foo")

        # verify
        self.assertEqual(oCells, ret)

    def test_get_named_cell(self):
        # prepare
        oCell = mock.Mock()
        oCells = mock.Mock()
        oCells.getCellByPosition.side_effect = [oCell]
        oRange = mock.Mock(ReferredCells=oCells)

        oRanges = mock.Mock()
        oRanges.getByName.side_effect = [oRange]
        oDoc = mock.Mock(NamedRanges=oRanges)

        # play
        ret = get_named_cell(oDoc, "foo")

        # verify
        self.assertEqual(oCell, ret)

    def test_get_main_cell(self):
        # prepare
        oMCell = mock.Mock()
        oCursor = mock.Mock(RangeAddress=mock.Mock(StartColumn=1, StartRow=5))
        oSheet = mock.Mock()
        oSheet.createCursorByRange.side_effect = [oCursor]
        oSheet.getCellByPosition.side_effect = [oMCell]
        oCell = mock.Mock(Spreadsheet=oSheet)

        # play
        ret = get_main_cell(oCell)

        # verify
        self.assertEqual(oMCell, ret)
        self.assertEqual([mock.call.collapseToMergedArea()],
                         oCursor.mock_calls)
        self.assertEqual([
            mock.call.createCursorByRange(oCell),
            mock.call.getCellByPosition(1, 5)
        ], oSheet.mock_calls)


##########################################################################
# STRUCTS
##########################################################################
class HelperStructTestCase(unittest.TestCase):
    @mock.patch("py4lo_helper.uno")
    def test_struct(self, uno):
        # prepare
        s = mock.Mock()
        uno.createUnoStruct.side_effect = [s]

        # play
        actual_s = create_uno_struct("uno.Struct", foo="bar")

        # verify
        self.assertIs(s, actual_s)
        self.assertEqual("bar", actual_s.foo)

    def test_make_pv(self):
        pv = make_pv("name", "value")
        self.assertEqual("com.sun.star.beans.PropertyValue", pv.typeName)
        self.assertEqual("name", pv.Name)
        self.assertEqual("value", pv.Value)

    def test_make_full_pv(self):
        pv = make_full_pv("name", "value", 20, PropertyState.AMBIGUOUS_VALUE)
        self.assertEqual("com.sun.star.beans.PropertyValue", pv.typeName)
        self.assertEqual("name", pv.Name)
        self.assertEqual("value", pv.Value)
        self.assertEqual(20, pv.Handle)
        self.assertEqual(PropertyState.AMBIGUOUS_VALUE, pv.State)

    def test_make_full_pv2(self):
        pv = make_full_pv("name", "value")
        self.assertTrue("com.sun.star.beans.PropertyValue", pv.typeName)
        self.assertEqual("name", pv.Name)
        self.assertEqual("value", pv.Value)
        self.assertEqual(0, pv.Handle)
        self.assertEqual(PropertyState.DIRECT_VALUE, pv.State)

    def test_make_pvs(self):
        pvs = make_pvs({"name1": "value1", "name2": "value2"})
        pvs = sorted(pvs, key=lambda pv: pv.Name)
        self.assertEqual(2, len(pvs))
        self.assertEqual("name1", pvs[0].Name)
        self.assertEqual("value1", pvs[0].Value)
        self.assertEqual("name2", pvs[1].Name)
        self.assertEqual("value2", pvs[1].Value)

    def test_make_pvs_none(self):
        pvs = make_pvs()
        self.assertEqual(0, len(pvs))

    def test_update_pvs(self):
        pvs = make_pvs({"name1": "value1", "name2": "value2"})
        update_pvs(pvs, {"name1": "value3"})
        pvs = sorted(pvs, key=lambda pv: pv.Name)
        self.assertEqual(2, len(pvs))
        self.assertEqual("name1", pvs[0].Name)
        self.assertEqual("value3", pvs[0].Value)
        self.assertEqual("name2", pvs[1].Name)
        self.assertEqual("value2", pvs[1].Value)

    def test_update_pvs_non_existing(self):
        pvs = make_pvs({"name1": "value1", "name2": "value2"})
        update_pvs(pvs, {"name3": "value3"})
        pvs = sorted(pvs, key=lambda pv: pv.Name)
        self.assertEqual(2, len(pvs))
        self.assertEqual("name1", pvs[0].Name)
        self.assertEqual("value1", pvs[0].Value)
        self.assertEqual("name2", pvs[1].Name)
        self.assertEqual("value2", pvs[1].Value)

    def test_make_locale(self):
        locale = make_locale("en", "US")
        self.assertEqual("com.sun.star.lang.Locale", locale.typeName)
        self.assertEqual("US", locale.Country)
        self.assertEqual("en", locale.Language)
        self.assertEqual("", locale.Variant)

    def test_make_locale_subtags(self):
        locale = make_locale("de", "CH", ["1996"])
        self.assertEqual("", locale.Country)
        self.assertEqual("qlt", locale.Language)
        self.assertEqual("de-CH-1996", locale.Variant)

    def test_make_locale_subtags_wo_region(self):
        locale = make_locale("sr", subtags=["Latn", "RS"])
        self.assertEqual("", locale.Country)
        self.assertEqual("qlt", locale.Language)
        self.assertEqual("sr-Latn-RS", locale.Variant)

    def test_make_border(self):
        border = make_border(0xFF0000, 3, BorderLineStyle.SOLID)
        self.assertEqual("com.sun.star.table.BorderLine2", border.typeName)
        self.assertEqual(16711680, border.Color)
        self.assertEqual(3, border.LineWidth)
        self.assertEqual(BorderLineStyle.SOLID, border.LineStyle)

    @mock.patch("py4lo_helper.uno")
    def test_make_sort_field(self, uno):
        # prepare
        struct = mock.Mock()
        uno.createUnoStruct.side_effect = [struct]

        # play
        make_sort_field(1, False)

        # verify
        self.assertEqual(1, struct.Field)
        self.assertFalse(struct.IsAscending)


#########################################################################
# RANGES
#########################################################################

class HelperRangesTestCase(unittest.TestCase):
    def setUp(self):
        self.xsc = mock.Mock()
        self.doc = mock.Mock()
        self.ctrl = mock.Mock()
        self.frame = mock.Mock()
        self.parent_win = mock.Mock()
        self.msp = mock.Mock()
        self.ctxt = mock.Mock()
        self.sm = mock.Mock()
        self.desktop = mock.Mock()
        py4lo_helper.provider = _ObjectProvider(
            self.doc, self.ctrl, self.frame, self.parent_win, self.msp,
            self.ctxt, self.sm, self.desktop)
        py4lo_helper._inspect = _Inspector(py4lo_helper.provider)

    def test_get_last_used_row(self):
        # prepare
        oCursor = mock.Mock(RangeAddress=mock.Mock(EndRow=10))
        oSheet = mock.Mock()
        oSheet.createCursor.side_effect = [oCursor]

        # play
        r = get_last_used_row(oSheet)

        # verify
        self.assertEqual(10, r)
        self.assertEqual([
            mock.call.gotoStartOfUsedArea(True),
            mock.call.gotoEndOfUsedArea(True)
        ], oCursor.mock_calls)

    def test_get_used_range_address(self):
        # prepare
        oRangeAddress = mock.Mock()
        oCursor = mock.Mock(RangeAddress=oRangeAddress)
        oSheet = mock.Mock()
        oSheet.createCursor.side_effect = [oCursor]

        # play
        ra = get_used_range_address(oSheet)

        # verify
        self.assertEqual(oRangeAddress, ra)
        self.assertEqual([
            mock.call.gotoStartOfUsedArea(True),
            mock.call.gotoEndOfUsedArea(True)
        ], oCursor.mock_calls)

    def test_used_range(self):
        # prepare
        oCursor = mock.Mock(RangeAddress=mock.Mock(
            StartColumn=1, StartRow=2, EndColumn=8, EndRow=10
        ))
        oRange = mock.Mock()
        oSheet = mock.Mock()
        oSheet.createCursor.side_effect = [oCursor]
        oSheet.getCellRangeByPosition.side_effect = [oRange]

        # play
        oActualRange = get_used_range(oSheet)

        # verify
        self.assertEqual(oRange, oActualRange)
        self.assertEqual([
            mock.call.gotoStartOfUsedArea(True),
            mock.call.gotoEndOfUsedArea(True)
        ], oCursor.mock_calls)
        self.assertEqual([
            mock.call.createCursor(),
            mock.call.getCellRangeByPosition(1, 2, 8, 10)
        ], oSheet.mock_calls)

    def test_used_range2(self):
        # prepare
        oRangeAddress = mock.Mock(
            StartColumn=2, StartRow=4, EndColumn=8, EndRow=16
        )
        oCursor = mock.Mock(RangeAddress=oRangeAddress)
        oSheet = mock.Mock()
        oSheet.createCursor.side_effect = [oCursor]
        oExpectedRange = mock.Mock()
        oSheet.getCellRangeByPosition.side_effect = [oExpectedRange]

        # play
        oRange = get_used_range(oSheet)

        # verify
        self.assertEqual([
            mock.call.gotoStartOfUsedArea(True),
            mock.call.gotoEndOfUsedArea(True)
        ], oCursor.mock_calls)
        self.assertEqual([
            mock.call.createCursor(),
            mock.call.getCellRangeByPosition(2, 4, 8, 16)
        ], oSheet.mock_calls)
        self.assertEqual(oExpectedRange, oRange)

    def test_narrow_range_to_address(self):
        # prepare
        oSheet = mock.Mock()
        oRangeAddress = mock.Mock(
            StartColumn=1, StartRow=2, EndColumn=4, EndRow=8
        )
        oExpectedRange = mock.Mock()
        oSheet.getCellRangeByPosition.side_effect = [oExpectedRange]

        # play
        oRange = narrow_range_to_address(oSheet, oRangeAddress)

        # verify
        self.assertEqual([mock.call.getCellRangeByPosition(1, 2, 4, 8)],
                         oSheet.mock_calls)
        self.assertEqual(oExpectedRange, oRange)

    def test_get_range_size(self):
        # prepare
        oRange = mock.Mock(RangeAddress=mock.Mock(
            StartColumn=1, StartRow=2, EndColumn=4, EndRow=8
        ))
        # play
        w, h = get_range_size(oRange)

        # verify
        self.assertEqual(4, w)
        self.assertEqual(7, h)

    @mock.patch("py4lo_helper.parent_doc")
    def test_copy_range(self, pd):
        # prepare
        oRange = mock.Mock()
        oController = mock.Mock()
        oDoc = mock.Mock(CurrentController=oController)
        pd.side_effect = [oDoc]
        oDisp = mock.Mock()
        self.sm.createInstance.side_effect = [oDisp]
        oRanges = mock.Mock()
        oDoc.createInstance.side_effect = [oRanges]

        # play
        copy_range(oRange)

        # verify
        self.assertEqual([
            mock.call.select(oRange),
            mock.call.select(oRanges),
        ], oController.mock_calls)
        self.assertEqual([
            mock.call.executeDispatch(oController, '.uno:Copy', '', 0, [])
        ], oDisp.mock_calls)

    @mock.patch("py4lo_helper.parent_doc")
    @mock.patch("py4lo_helper.make_pv")
    def test_paste_range(self, mkpv, pd):
        mkpv.side_effect = lambda a, b: (a, b)

        # prepare
        oSheet = mock.Mock()
        oDestAddress = mock.Mock()
        oController = mock.Mock()
        oDoc = mock.Mock(CurrentController=oController)
        oRange = mock.Mock()
        oRanges = mock.Mock()
        oSheet.getCellByPosition.side_effect = [oRange]
        oDoc.createInstance.side_effect = [oRanges]
        pd.side_effect = [oDoc]
        oDisp = mock.Mock()
        self.sm.createInstance.side_effect = [oDisp]

        # play
        paste_range(oSheet, oDestAddress)

        # verify
        self.assertEqual([
            mock.call.select(oRange),
            mock.call.select(oRanges),
        ], oController.mock_calls)
        self.assertEqual([
            mock.call.executeDispatch(
                oController, '.uno:InsertContents', '', 0,
                make_pvs({
                    "Flags": "SVDT", "FormulaCommand": 0,
                    "SkipEmptyCells": False,
                    "Transpose": False, "AsLink": False,
                    "MoveMode": 4
                }))
        ], oDisp.mock_calls)

    @mock.patch("py4lo_helper.parent_doc")
    def test_paste_range_formulas(self, pd):
        # prepare
        oSheet = mock.Mock()
        oDestAddress = mock.Mock()
        oController = mock.Mock()
        oDoc = mock.Mock(CurrentController=oController)
        oRange = mock.Mock()
        oRanges = mock.Mock()
        oSheet.getCellByPosition.side_effect = [oRange]
        oDoc.createInstance.side_effect = [oRanges]
        pd.side_effect = [oDoc]
        oDisp = mock.Mock()
        self.sm.createInstance.side_effect = [oDisp]

        # play
        paste_range(oSheet, oDestAddress, True)

        # verify
        self.assertEqual([
            mock.call.select(oRange),
            mock.call.select(oRanges),
        ], oController.mock_calls)
        self.assertEqual([
            mock.call.executeDispatch(oController, '.uno:Paste', '', 0, [])
        ], oDisp.mock_calls)

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


##########################################################################
# DATA ARRAY
##########################################################################
class HelperDataArrayTestCase(unittest.TestCase):
    def test_row_count(self):
        data_array = [
            ("", "", "", "", "", "", "",),
            ("", "", "a", "", "", "b", "",),
            ("", "", "", "oTextRange", "d", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "e", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",)
        ]
        self.assertEqual(1, top_void_row_count(data_array))
        self.assertEqual(2, bottom_void_row_count(data_array))
        self.assertEqual(2, left_void_row_count(data_array))
        self.assertEqual(1, right_void_row_count(data_array))

    def test_row_count_no_row(self):
        data_array = []
        self.assertEqual(0, top_void_row_count(data_array))
        self.assertEqual(0, bottom_void_row_count(data_array))
        self.assertEqual(0, left_void_row_count(data_array))
        self.assertEqual(0, right_void_row_count(data_array))

    def test_row_count_one_cell(self):
        data_array = [("",)]
        self.assertEqual(1, top_void_row_count(data_array))
        self.assertEqual(1, bottom_void_row_count(data_array))
        self.assertEqual(1, left_void_row_count(data_array))
        self.assertEqual(1, right_void_row_count(data_array))

    def test_row_count_one_row(self):
        data_array = [("", "")]
        self.assertEqual(1, top_void_row_count(data_array))
        self.assertEqual(1, bottom_void_row_count(data_array))
        self.assertEqual(2, left_void_row_count(data_array))
        self.assertEqual(2, right_void_row_count(data_array))

    def test_row_count_one_col(self):
        data_array = [("",), ("",)]
        self.assertEqual(2, top_void_row_count(data_array))
        self.assertEqual(2, bottom_void_row_count(data_array))
        self.assertEqual(1, left_void_row_count(data_array))
        self.assertEqual(1, right_void_row_count(data_array))

    def test_row_count2(self):
        data_array = [
            ("", "", "", "", "", "", "",),
            ("", "", "a", "", "", "b", "",),
            ("", "", "", "oTextRange", "d", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "e", "", "",),
            ("", "", "", "", "", "", "",),
            ("", "", "", "", "", "", "",)
        ]
        self.assertEqual(1, top_void_row_count(data_array))
        self.assertEqual(2, bottom_void_row_count(data_array))
        self.assertEqual(2, left_void_row_count(data_array))
        self.assertEqual(1, right_void_row_count(data_array))

    def test_row_count_empty(self):
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

    @mock.patch("py4lo_helper.get_used_range")
    def test_data_array(self, gura):
        # prepare
        oSheet = mock.Mock()
        da = [("foo",), ("bar",), ("baz",)]
        oRange = mock.Mock(DataArray=da)
        gura.side_effect = [oRange]
        # play
        act_da = data_array(oSheet)

        # verify
        self.assertEqual(da, act_da)

    @mock.patch("py4lo_helper.get_used_range")
    def test_data_array_no_row(self, gura):
        # prepare
        oSheet = mock.Mock()
        da = []
        oRange = mock.Mock(DataArray=da)
        gura.side_effect = [oRange]
        # play
        act_da = data_array(oSheet)

        # verify
        self.assertEqual(da, act_da)


##########################################################################
# FORMATTING
##########################################################################

class HelperFormattingTestCase(unittest.TestCase):
    def setUp(self):
        self.xsc = mock.Mock()
        self.doc = mock.Mock()
        self.ctrl = mock.Mock()
        self.frame = mock.Mock()
        self.parent_win = mock.Mock()
        self.msp = mock.Mock()
        self.ctxt = mock.Mock()
        self.sm = mock.Mock()
        self.desktop = mock.Mock()
        py4lo_helper.provider = _ObjectProvider(
            self.doc, self.ctrl, self.frame, self.parent_win, self.msp,
            self.ctxt, self.sm, self.desktop)
        py4lo_helper._inspect = _Inspector(py4lo_helper.provider)

    def test_validation_builder(self):
        # prepare
        oVal = mock.Mock()
        oCell = mock.Mock(Validation=oVal)

        # play
        builder = ListValidationBuilder()
        builder.values(["1", "2", "3"]).default_string("1").ignore_blank(
            True).sort_values(True).show_error(False)
        builder.update(oCell)

        # verify
        self.assertEqual('"1";"2";"3"', oVal.Formula1)
        self.assertTrue(oVal.IgnoreBlankCells)
        self.assertFalse(oVal.ShowErrorMessage)
        self.assertEqual(1, oVal.ShowList)
        self.assertEqual(ValidationType.LIST, oVal.Type)

    def test_set_validation_list(self):
        # prepare
        oVal = mock.Mock()
        oCell = mock.Mock(Validation=oVal)

        # play
        set_validation_list_by_cell(oCell, ["a", "b", "oTextRange"])

        # verify
        self.assertEqual(ValidationType.LIST, oVal.Type)
        self.assertEqual(2, oVal.ShowList)
        self.assertTrue(oVal.ShowErrorMessage)
        self.assertTrue(oVal.IgnoreBlankCells)
        self.assertEqual('"a";"b";"oTextRange"', oVal.Formula1)
        self.assertIsNotNone(oCell.String)

    def test_set_validation_list2(self):
        # prepare
        oVal = mock.Mock()
        oCell = mock.Mock(Validation=oVal)

        # play
        set_validation_list_by_cell(oCell, [1, 2, 3], "foo", False, True,
                                    False)

        # verify
        self.assertEqual(ValidationType.LIST, oVal.Type)
        self.assertEqual(1, oVal.ShowList)
        self.assertFalse(oVal.ShowErrorMessage)
        self.assertFalse(oVal.IgnoreBlankCells)
        self.assertEqual('"1";"2";"3"', oVal.Formula1)
        self.assertEqual("foo", oCell.String)

    def test_set_validation_list3(self):
        # prepare
        oVal = mock.Mock()
        oCell = mock.Mock(Validation=oVal)

        # play
        set_validation_list_by_cell(oCell, [1, "a", False])

        # verify
        self.assertEqual(ValidationType.LIST, oVal.Type)
        self.assertEqual(2, oVal.ShowList)
        self.assertTrue(oVal.ShowErrorMessage)
        self.assertTrue(oVal.IgnoreBlankCells)
        self.assertEqual('"1";"a";"False"', oVal.Formula1)
        self.assertIsNotNone(oCell.String)

    @mock.patch("py4lo_helper.update_pvs")
    def test_sort_range(self, up):
        # prepare
        oSD = mock.Mock()
        oRange = mock.Mock()
        oRange.createSortDescriptor.side_effect = [oSD]
        sort_fields = tuple([mock.Mock(), mock.Mock()])

        # play
        sort_range(oRange, sort_fields, True)

        # verify
        self.assertEqual([
            mock.call.createSortDescriptor(),
            mock.call.sort(oSD)
        ], oRange.mock_calls)
        self.assertEqual([
            mock.call(oSD, {'ContainsHeader': True, 'SortFields': mock.ANY})
        ], up.mock_calls)

    @mock.patch("py4lo_helper.update_pvs")
    def test_sort_range_no_header(self, up):
        # prepare
        oSD = mock.Mock()
        oRange = mock.Mock()
        oRange.createSortDescriptor.side_effect = [oSD]
        sort_fields = tuple([mock.Mock(), mock.Mock()])

        # play
        sort_range(oRange, sort_fields, False)

        # verify
        self.assertEqual([
            mock.call.createSortDescriptor(),
            mock.call.sort(oSD)
        ], oRange.mock_calls)
        self.assertEqual([
            mock.call(oSD, {'ContainsHeader': False, 'SortFields': mock.ANY})
        ], up.mock_calls)

    def test_quote_element(self):
        self.assertEqual('"A"', quote_element("A"))
        self.assertEqual('"1.5"', quote_element(1.5))
        self.assertEqual('"True"', quote_element(True))

    def test_clear_conditional_format(self):
        # prepare
        oFormat = mock.Mock()
        oRange = mock.Mock(ConditionalFormat=oFormat)

        # play
        clear_conditional_format(oRange)

        # verify
        self.assertEqual([mock.call.clear()], oFormat.mock_calls)

    @mock.patch("py4lo_helper.get_formula_conditional_entry")
    def test_conditional_format_on_formulas(self, gfce):
        # prepare
        oFormat = mock.Mock()
        oRange = mock.Mock(ConditionalFormat=oFormat)
        oCellAddress = mock.Mock()
        gfce.side_effect = ["e"]

        # play
        conditional_format_on_formulas(
            oRange, {"foo": "bar"}, oCellAddress)

        # verify
        self.assertEqual([mock.call('foo', 'bar', oCellAddress)],
                         gfce.mock_calls)
        self.assertEqual([mock.call.addNew("e")], oFormat.mock_calls)

    @mock.patch("py4lo_helper.make_pv")
    def test_get_formula_conditional_entry(self, mk_pv):
        # prepare
        oCellAddress = mock.Mock()
        mk_pv.side_effect = [1, 2, 3, 4, 5]

        # play
        ret = get_formula_conditional_entry("A1=0", "foo", oCellAddress)

        # verify
        self.assertEqual((1, 2, 3, 4, 5), ret)
        self.assertEqual([
            mock.call('Formula1', 'A1=0'),
            mock.call('Formula2', ''),
            mock.call('Operator', ConditionOperator.FORMULA),
            mock.call('StyleName', 'foo'),
            mock.call('SourcePosition', oCellAddress)
        ], mk_pv.mock_calls)

    def test_find_number_format_style(self):
        # prepare
        oFormats = mock.Mock()
        oFormats.queryKey.side_effect = [9]
        oDoc = mock.Mock(NumberFormats=oFormats)
        oLocale = mock.Mock()

        # play
        ret = find_or_create_number_format_style(oDoc, "YY-MM-DD", oLocale)

        # verify
        self.assertEqual(9, ret)

    @mock.patch("py4lo_helper.make_locale")
    def test_find_number_format_style_no_locale(self, ml):
        # prepare
        oFormats = mock.Mock()
        oFormats.queryKey.side_effect = [9]
        oDoc = mock.Mock(NumberFormats=oFormats)
        oLocale = mock.Mock()
        ml.side_effect = [oLocale]

        # play
        ret = find_or_create_number_format_style(oDoc, "YY-MM-DD")

        # verify
        self.assertEqual(9, ret)

    def test_create_number_format_style(self):
        # prepare
        oFormats = mock.Mock()
        oFormats.queryKey.side_effect = [-1]
        oFormats.addNew.side_effect = [10]
        oDoc = mock.Mock(NumberFormats=oFormats)
        oLocale = mock.Mock()

        # play
        ret = find_or_create_number_format_style(oDoc, "YY-MM-DD", oLocale)

        # verify
        self.assertEqual(10, ret)

    @mock.patch("py4lo_helper.parent_doc")
    def test_create_filter(self, pd):
        # prepare
        oController = mock.Mock()
        oDoc = mock.Mock(CurrentController=oController)
        oRange = mock.Mock()
        pd.side_effect = [oDoc]

        # play
        create_filter(oRange)

        # verify
        self.assertEqual([
            mock.call.select(oRange),
            mock.call.select(mock.ANY)
        ], oController.mock_calls)
        self.assertEqual([
            mock.call.createInstance('com.sun.star.frame.DispatchHelper'),
            mock.call.createInstance().executeDispatch(
                oController, '.uno:DataFilterAutoFilter', '', 0, [])
        ], self.sm.mock_calls)

    def test_row_as_header(self):
        # prepare
        oRow = mock.Mock()

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
        oColumn = mock.Mock(Width=1 * 1000)

        # play
        column_optimal_width(oColumn)

        # verify
        self.assertFalse(oColumn.OptimalWidth)
        self.assertEqual(2 * 1000, oColumn.Width)

    def test_column_optimal_width_large(self):
        # prepare
        oColumn = mock.Mock(Width=20 * 1000)

        # play
        column_optimal_width(oColumn)

        # verify
        self.assertFalse(oColumn.OptimalWidth)
        self.assertTrue(oColumn.IsTextWrapped)
        self.assertEqual(10 * 1000, oColumn.Width)

    def test_column_optimal_width_medium(self):
        # prepare
        oColumn = mock.Mock(Width=5 * 1000)

        # play
        column_optimal_width(oColumn)

        # verify
        self.assertTrue(oColumn.OptimalWidth)

    @mock.patch("py4lo_helper.get_used_range_address")
    def test_set_print_area(self, gura):
        # prepare
        sheet_ra = mock.Mock()
        gura.side_effect = [sheet_ra]
        oSheet = mock.Mock()
        title_ra = mock.Mock()
        oRow = mock.Mock(RangeAddress=title_ra)

        # play
        set_print_area(oSheet, oRow)

        # verify
        self.assertEqual([
            mock.call.setPrintAreas([sheet_ra]),
            mock.call.setPrintTitleRows(True),
            mock.call.setTitleRows(title_ra)
        ], oSheet.mock_calls)

    @mock.patch("py4lo_helper.get_used_range_address")
    def test_set_print_area_no_title(self, gura):
        # prepare
        sheet_ra = mock.Mock()
        gura.side_effect = [sheet_ra]
        oSheet = mock.Mock()

        # play
        set_print_area(oSheet)

        # verify
        self.assertEqual([
            mock.call.setPrintAreas([sheet_ra])
        ], oSheet.mock_calls)

    @mock.patch("py4lo_helper.parent_doc")
    def test_get_page_style(self, pd):
        # prepare
        oSheet = mock.Mock(PageStyle="foo")
        oDoc = mock.Mock()
        pd.side_effect = [oDoc]

        # play
        _ = get_page_style(oSheet)

        # verify
        self.assertEqual([
            mock.call.StyleFamilies.getByName('PageStyles'),
            mock.call.StyleFamilies.getByName().getByName('foo')
        ], oDoc.mock_calls)

    @mock.patch("py4lo_helper.create_uno_struct")
    @mock.patch("py4lo_helper.get_used_range")
    @mock.patch("py4lo_helper.get_page_style")
    def test_set_paper_1(self, gps, gur, ms):
        self._set_paper(gps, gur, ms, 10 * 1000, 40 * 1000, False, 1, 0, 21000,
                        29700)

    @mock.patch("py4lo_helper.create_uno_struct")
    @mock.patch("py4lo_helper.get_used_range")
    @mock.patch("py4lo_helper.get_page_style")
    def test_set_paper_2(self, gps, gur, ms):
        self._set_paper(gps, gur, ms, 40 * 1000, 400 * 1000, False, 1, 0,
                        29700, 42000)

    @mock.patch("py4lo_helper.create_uno_struct")
    @mock.patch("py4lo_helper.get_used_range")
    @mock.patch("py4lo_helper.get_page_style")
    def test_set_paper_3(self, gps, gur, ms):
        self._set_paper(gps, gur, ms, 40 * 1000, 10 * 1000, True, 0, 1, 29700,
                        21000)

    @mock.patch("py4lo_helper.create_uno_struct")
    @mock.patch("py4lo_helper.get_used_range")
    @mock.patch("py4lo_helper.get_page_style")
    def test_set_paper_4(self, gps, gur, ms):
        self._set_paper(gps, gur, ms, 400 * 1000, 40 * 1000, True, 0, 1, 42000,
                        29700)

    def _set_paper(self, gps, gur, ms, w, h, is_landscape, sx, sy, pw, ph):
        # prepare
        oSheet = mock.Mock()
        oPageStyle = mock.Mock()
        gps.side_effect = [oPageStyle]
        oRange = mock.Mock(Size=mock.Mock(Width=w, Height=h))
        gur.side_effect = [oRange]
        # play
        set_paper(oSheet)
        # verify
        self.assertEqual(is_landscape, oPageStyle.IsLandscape)
        self.assertEqual(sx, oPageStyle.ScaleToPagesX)
        self.assertEqual(sy, oPageStyle.ScaleToPagesY)
        self.assertEqual([
            mock.call('com.sun.star.awt.Size', Width=pw, Height=ph)],
            ms.mock_calls)

    @mock.patch("py4lo_helper.parent_doc")
    def test_add_link(self, pd):
        # prepare
        oText = mock.Mock(End=1)
        oCur = mock.Mock()
        oText.createTextCursorByRange.side_effect = [oCur]
        oCell = mock.Mock(Text=oText)
        oDoc = mock.Mock()
        oURL = mock.Mock()
        oDoc.createInstance.side_effect = [oURL]
        pd.side_effect = [oDoc]

        # play
        add_link(oCell, "a very long text", "https://url.org", -1)

        # verify
        self.assertEqual([
            mock.call.Text.createTextCursorByRange(1),
            mock.call.insertTextContent(oCur, oURL, False)
        ], oCell.mock_calls)
        self.assertEqual('a very long text', oURL.Representation)
        self.assertEqual("https://url.org", oURL.URL)

    @mock.patch("py4lo_helper.parent_doc")
    def test_add_link_wrapped(self, pd):
        # prepare
        oText = mock.Mock(End=1)
        oCur = mock.Mock()
        oText.createTextCursorByRange.side_effect = [oCur]
        oCell = mock.Mock(Text=oText)
        oDoc = mock.Mock()
        oURL1 = mock.Mock()
        oURL2 = mock.Mock()
        oDoc.createInstance.side_effect = [oURL1, oURL2]
        pd.side_effect = [oDoc]

        # play
        add_link(oCell, "a very long text", "https://url.org", 10)

        # verify
        self.assertEqual([
            mock.call.Text.createTextCursorByRange(1),
            mock.call.insertTextContent(oCur, oURL1, False),
            mock.call.insertTextContent(oCur, oURL2, False)
        ], oCell.mock_calls)
        self.assertEqual('a very long', oURL1.Representation)
        self.assertEqual("https://url.org", oURL1.URL)
        self.assertEqual('text', oURL2.Representation)
        self.assertEqual("https://url.org", oURL2.URL)

    def test_wrap_text(self):
        self.assertEqual(['abcd', 'efgh', 'ij'],
                         _wrap_text("abcd efgh ij", 6))
        self.assertEqual(['abcd', 'efgh ij'],
                         _wrap_text("abcd efgh ij", 7))
        self.assertEqual(['abcd', 'efgh ij'],
                         _wrap_text("abcd efgh ij", 8))
        self.assertEqual(['abcd efgh', 'ij'],
                         _wrap_text("abcd efgh ij", 9))
        self.assertEqual(['abcd efgh', 'ij'],
                         _wrap_text("abcd efgh ij", 10))
        self.assertEqual(['abcd efgh ij'],
                         _wrap_text("abcd efgh ij", 11))
        self.assertEqual(['a very', 'long', 'text'],
                         _wrap_text("a very long text", 8))
        self.assertEqual(['a very', 'long text'],
                         _wrap_text("a very long text", 9))


###########################################################################
# OPEN A DOCUMENT
###########################################################################

class HelperOpenTestCase(unittest.TestCase):
    def setUp(self):
        self.xsc = mock.Mock()
        self.doc = mock.Mock()
        self.ctrl = mock.Mock()
        self.frame = mock.Mock()
        self.parent_win = mock.Mock()
        self.msp = mock.Mock()
        self.ctxt = mock.Mock()
        self.sm = mock.Mock()
        self.desktop = mock.Mock()
        py4lo_helper.provider = _ObjectProvider(
            self.doc, self.ctrl, self.frame, self.parent_win, self.msp,
            self.ctxt, self.sm, self.desktop)
        py4lo_helper._inspect = _Inspector(py4lo_helper.provider)


    @mock.patch("py4lo_helper.make_pv")
    def test_open_in_calc(self, mkpv):
        mkpv.side_effect = lambda a, b: (a, b)

        # prepare
        py4lo_helper.provider = mock.Mock()

        # play
        open_in_calc("/fname", Target.SELF, FrameSearchFlag.AUTO, Hidden=True)
        # verify

        self.assertEqual([
            mock.call.desktop.loadComponentFromURL(
                'file:///fname',
                Target.SELF, 0, (("Hidden", True),))],
            py4lo_helper.provider.mock_calls)

    def test_open_in_calc_no_params(self):
        # prepare
        py4lo_helper.provider = mock.Mock()

        # play
        open_in_calc("/fname", Target.SELF, FrameSearchFlag.AUTO)
        # verify

        self.assertEqual([
            mock.call.desktop.loadComponentFromURL(
                'file:///fname',
                Target.SELF, 0, tuple())],
            py4lo_helper.provider.mock_calls)

    def test_new_doc(self):
        # prepare
        oDoc = mock.Mock()
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        oActualDoc = new_doc(NewDocumentUrl.Calc)

        # verify
        self.assertEqual(oDoc, oActualDoc)
        self.assertEqual([
            mock.call.loadComponentFromURL(
                'private:factory/scalc', Target.BLANK, 0, tuple())
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers, mock.call.unlockControllers
        ], oDoc.mock_calls)

    def test_doc_builder(self):
        # prepare
        oDoc = mock.Mock()
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        d = doc_builder(NewDocumentUrl.Calc)
        d.build()

        # verify
        self.assertEqual([
            mock.call.loadComponentFromURL(
                'private:factory/scalc', Target.BLANK, 0, tuple())
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers, mock.call.unlockControllers
        ], oDoc.mock_calls)

    def test_doc_builder_pvs(self):
        # prepare
        oDoc = mock.Mock()
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        d = doc_builder(NewDocumentUrl.Calc, pvs={"Hidden": True})
        d.build()

        # verify
        self.assertEqual([
            mock.call.loadComponentFromURL('private:factory/scalc',
                                           Target.BLANK, 0,
                                           {"Hidden": True})
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers, mock.call.unlockControllers
        ], oDoc.mock_calls)

    def test_doc_builder_sheet_names_two(self):
        # prepare
        sheets = [mock.Mock() for _ in range(3)]
        oSheets = mock.Mock(Count=3)
        oSheets.getByIndex.side_effect = lambda i: sheets[i]
        oDoc = mock.Mock(Sheets=oSheets)
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        d = doc_builder(NewDocumentUrl.Calc)
        d.sheet_names(list("ab"), expand_if_necessary=False,
                      trunc_if_necessary=False)
        d.build()

        # verify
        self.assertEqual([
            mock.call.loadComponentFromURL(
                NewDocumentUrl.Calc, Target.BLANK, 0, ())
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers(),
            mock.call.Sheets.getByIndex(0),
            mock.call.Sheets.getByIndex(1),
            mock.call.Sheets.getByIndex(2),
            mock.call.unlockControllers()
        ], oDoc.mock_calls)
        self.assertEqual('a', sheets[0].Name)
        self.assertEqual('b', sheets[1].Name)
        self.assertEqual([], sheets[2].mock_calls)

    def test_doc_builder_sheet_names_two_trunc(self):
        # prepare
        sheets = [mock.Mock() for _ in range(3)]
        sheets[2].Name = "foo"
        oSheets = mock.MagicMock()
        type(oSheets).Count = mock.PropertyMock(side_effect=[3, 3, 2])
        oSheets.getByIndex.side_effect = lambda i: sheets[i]
        oDoc = mock.Mock(Sheets=oSheets)
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        d = doc_builder(NewDocumentUrl.Calc)
        d.sheet_names(list("ab"), expand_if_necessary=False,
                      trunc_if_necessary=True)
        d.build()

        # verify
        self.assertEqual([
            mock.call.loadComponentFromURL(
                NewDocumentUrl.Calc, Target.BLANK, 0, ())
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers(),
            mock.call.Sheets.getByIndex(0),
            mock.call.Sheets.getByIndex(1),
            mock.call.Sheets.getByIndex(2),
            mock.call.Sheets.getByIndex(2),
            mock.call.Sheets.removeByName('foo'),
            mock.call.unlockControllers()
        ], oDoc.mock_calls)
        self.assertEqual('a', sheets[0].Name)
        self.assertEqual('b', sheets[1].Name)
        self.assertEqual([], sheets[2].mock_calls)

    def test_doc_builder_sheet_names_three(self):
        # prepare
        sheets = [mock.Mock() for _ in range(3)]
        oSheets = mock.Mock(Count=3)
        oSheets.getByIndex.side_effect = lambda i: sheets[i]
        oDoc = mock.Mock(Sheets=oSheets)
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        d = doc_builder(NewDocumentUrl.Calc)
        d.sheet_names(list("abc"), expand_if_necessary=False,
                      trunc_if_necessary=False)
        d.build()

        # verify
        self.assertEqual([
            mock.call.loadComponentFromURL(
                NewDocumentUrl.Calc, Target.BLANK, 0, ())
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers(),
            mock.call.Sheets.getByIndex(0),
            mock.call.Sheets.getByIndex(1),
            mock.call.Sheets.getByIndex(2),
            mock.call.unlockControllers()
        ], oDoc.mock_calls)
        self.assertEqual('a', sheets[0].Name)
        self.assertEqual('b', sheets[1].Name)
        self.assertEqual('c', sheets[2].Name)

    def test_doc_builder_sheet_names_four(self):
        # prepare
        sheets = [mock.Mock() for _ in range(3)]
        oSheets = mock.Mock(Count=3)
        oSheets.getByIndex.side_effect = lambda i: sheets[i]
        oDoc = mock.Mock(Sheets=oSheets)
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        d = doc_builder(NewDocumentUrl.Calc)
        d.sheet_names(list("abcc"), expand_if_necessary=True)
        d.build()

        # verify
        self.assertEqual([
            mock.call.loadComponentFromURL(
                NewDocumentUrl.Calc, Target.BLANK, 0, ())
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers(),
            mock.call.Sheets.getByIndex(0),
            mock.call.Sheets.getByIndex(1),
            mock.call.Sheets.getByIndex(2),
            mock.call.Sheets.insertNewByName('c', 3),
            mock.call.unlockControllers()
        ], oDoc.mock_calls)
        self.assertEqual('a', sheets[0].Name)
        self.assertEqual('b', sheets[1].Name)
        self.assertEqual('c', sheets[2].Name)

    def test_doc_builder_apply(self):
        # prepare
        sheets = [mock.Mock() for _ in range(3)]
        oSheets = mock.Mock(Count=3)
        oSheets.Types = [mock.Mock(typeName="com.sun.star.container.XIndexAccess")]
        oSheets.getByIndex.side_effect = lambda i: sheets[i]
        oDoc = mock.Mock(Sheets=oSheets)
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        d = doc_builder(NewDocumentUrl.Calc)

        def func(oSheet): oSheet.app = 0

        d.apply_func_to_sheets(func)
        d.build()

        # verify
        self.assertEqual([
            mock.call.loadComponentFromURL(
                NewDocumentUrl.Calc, Target.BLANK, 0, ())
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers(),
            mock.call.Sheets.getByIndex(0),
            mock.call.Sheets.getByIndex(1),
            mock.call.Sheets.getByIndex(2),
            mock.call.unlockControllers()
        ], oDoc.mock_calls)
        for s in sheets:
            self.assertEqual([], s.mock_calls)
            self.assertEqual(0, s.app)

    def test_doc_builder_apply_list(self):
        # prepare
        sheets = [mock.Mock(app=None) for _ in range(3)]
        oSheets = mock.Mock(Count=3)
        oSheets.Types = [mock.Mock(typeName="com.sun.star.container.XIndexAccess")]
        oSheets.getByIndex.side_effect = lambda i: sheets[i]
        oDoc = mock.Mock(Sheets=oSheets)
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        d = doc_builder(NewDocumentUrl.Calc)

        def func1(oSheet): oSheet.app = 0

        def func2(oSheet): oSheet.app = 1

        d.apply_func_list_to_sheets([func1, func2])
        d.build()

        # verify
        self.assertEqual([
            mock.call.loadComponentFromURL(
                NewDocumentUrl.Calc, Target.BLANK, 0, ())
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers(),
            mock.call.Sheets.getByIndex(0),
            mock.call.Sheets.getByIndex(1),
            mock.call.unlockControllers()
        ], oDoc.mock_calls)
        for s in sheets:
            self.assertEqual([], s.mock_calls)

        self.assertEqual(0, sheets[0].app)
        self.assertEqual(1, sheets[1].app)
        self.assertIsNone(sheets[2].app)

    def test_doc_builder_duplicate_base_sheet(self):
        # prepare
        sheets = [mock.Mock(Name=str(i)) for i in range(3)]
        oSheets = mock.Mock()
        oSheets.getCount.side_effect = [3]
        oSheets.getByIndex.side_effect = lambda i: sheets[i]
        oDoc = mock.Mock(Sheets=oSheets)
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        d = doc_builder(NewDocumentUrl.Calc)

        def func(oSheet): oSheet.app = 0

        d.duplicate_base_sheet(func, list("abc"))
        d.build()

        # verify
        self.assertEqual([
            mock.call.loadComponentFromURL(
                NewDocumentUrl.Calc, Target.BLANK, 0, ())
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers(),
            mock.call.Sheets.getByIndex(0),
            mock.call.Sheets.copyByName('0', 'a', 1),
            mock.call.Sheets.copyByName('0', 'b', 2),
            mock.call.Sheets.copyByName('0', 'c', 3),
            mock.call.unlockControllers()
        ], oDoc.mock_calls)
        for s in sheets:
            self.assertEqual([], s.mock_calls)

        self.assertEqual(0, sheets[0].app)
        self.assertNotEqual(0, sheets[1].app)
        self.assertNotEqual(0, sheets[2].app)

    def test_doc_builder_duplicate_to(self):
        # prepare
        sheets = [mock.Mock(Name=str(i)) for i in range(3)]
        oSheets = mock.Mock()
        oSheets.getCount.side_effect = [3]
        oSheets.getByIndex.side_effect = lambda i: sheets[i]
        oDoc = mock.Mock(Sheets=oSheets)
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        d = doc_builder(NewDocumentUrl.Calc)
        d.duplicate_to(6)
        d.build()

        # verify
        self.assertEqual([
            mock.call.loadComponentFromURL(
                NewDocumentUrl.Calc, Target.BLANK, 0, ())
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers(),
            mock.call.Sheets.getByIndex(0),
            mock.call.Sheets.copyByName('0', '00', 0),
            mock.call.Sheets.copyByName('0', '01', 1),
            mock.call.Sheets.copyByName('0', '02', 2),
            mock.call.Sheets.copyByName('0', '03', 3),
            mock.call.Sheets.copyByName('0', '04', 4),
            mock.call.Sheets.copyByName('0', '05', 5),
            mock.call.Sheets.copyByName('0', '06', 6),
            mock.call.unlockControllers()
        ], oDoc.mock_calls)
        for s in sheets:
            self.assertEqual([], s.mock_calls)

    def test_doc_builder_make_base_sheet(self):
        # prepare
        sheets = [mock.Mock(Name=str(i)) for i in range(3)]
        oSheets = mock.Mock()
        oSheets.getCount.side_effect = [3]
        oSheets.getByIndex.side_effect = lambda i: sheets[i]
        oDoc = mock.Mock(Sheets=oSheets)
        py4lo_helper.provider.desktop.loadComponentFromURL.side_effect = [oDoc]

        # play
        d = doc_builder(NewDocumentUrl.Calc)

        def func(oSheet): oSheet.app = 0

        d.make_base_sheet(func)
        d.build()

        # verify
        self.assertEqual([
            mock.call.loadComponentFromURL(
                NewDocumentUrl.Calc, Target.BLANK, 0, ())
        ], py4lo_helper.provider.desktop.mock_calls)
        self.assertEqual([
            mock.call.lockControllers(),
            mock.call.Sheets.getByIndex(0),
            mock.call.unlockControllers()
        ], oDoc.mock_calls)
        for s in sheets:
            self.assertEqual([], s.mock_calls)

        self.assertEqual(0, sheets[0].app)
        self.assertNotEqual(0, sheets[1].app)
        self.assertNotEqual(0, sheets[2].app)


##########################################################################
# MISC
##########################################################################


class HelperMiscTestCase(unittest.TestCase):
    def setUp(self):
        self.xsc = mock.Mock()
        self.doc = mock.Mock()
        self.ctrl = mock.Mock()
        self.frame = mock.Mock()
        self.parent_win = mock.Mock()
        self.msp = mock.Mock()
        self.ctxt = mock.Mock()
        self.sm = mock.Mock()
        self.desktop = mock.Mock()
        py4lo_helper.provider = _ObjectProvider(
            self.doc, self.ctrl, self.frame, self.parent_win, self.msp,
            self.ctxt, self.sm, self.desktop)
        py4lo_helper._inspect = _Inspector(py4lo_helper.provider)

    def testUnoService(self):
        create_uno_service_ctxt("x")
        self.sm.createInstanceWithContext.assert_called_once_with("x",
                                                                  self.ctxt)

        create_uno_service_ctxt("y", [1, 2, 3])
        self.sm.createInstanceWithArgumentsAndContext.assert_called_once_with(
            "y", [1, 2, 3], self.ctxt)

        create_uno_service("z")
        self.sm.createInstance.assert_called_once_with("z")

    def test_read_options(self):
        # prepare
        oRange = mock.Mock(DataArray=[
            ("foo", 1, "", ""),
            ("bar", 1, 2, 3),
            ("baz", "a", "", "b"),
            ("", "x", "", ""),
        ])
        oSheet = mock.Mock()
        oSheet.getCellRangeByPosition.side_effect = [oRange]
        aAddress = mock.Mock()
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
        oSheet = mock.Mock()
        aAddress = mock.Mock()
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

    @mock.patch("py4lo_helper.read_options")
    @mock.patch("py4lo_helper.get_used_range_address")
    def test_read_options_from_sheet_name(self, gura, ro):
        # prepare
        oSheet = mock.Mock()
        oSheets = mock.Mock()
        oSheets.getByName.side_effect = [oSheet]
        oDoc = mock.Mock(Sheets=oSheets)
        oAddress = mock.Mock()
        gura.side_effect = [oAddress]

        def f(k, v): return k, v

        # play
        read_options_from_sheet_name(oDoc, "foo", f)

        # verify
        self.assertEqual([mock.call(oSheet, oAddress, f)], ro.mock_calls)

    def test_copy_row_at_index(self):
        # prepare
        oRange = mock.Mock()
        oSheet = mock.Mock()
        oSheet.getCellRangeByPosition.side_effect = [oRange]
        row = ("foo", "bar", "baz")

        # play
        copy_row_at_index(oSheet, row, 3)

        # verifiy
        self.assertEqual([mock.call.getCellRangeByPosition(0, 3, 2, 3)],
                         oSheet.mock_calls)
        self.assertEqual([row], oRange.DataArray)


class MiscTestCase(unittest.TestCase):
    def setUp(self):
        self.oSP = mock.Mock()
        self.provider = mock.Mock(script_provider=self.oSP)
        self.inspector = _Inspector(self.provider)

    def test_use_xray_ok(self):
        # prepare
        # play
        self.inspector.use_xray(False)

        # verify
        self.assertEqual([mock.call.getScript(
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
        self.assertEqual([mock.call.getScript(
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
        self.assertEqual([mock.call.getScript(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')],
            self.oSP.mock_calls)
        self.assertTrue(self.inspector._ignore_xray)

    def test_xray(self):
        # prepare
        obj = mock.Mock()

        # play
        self.inspector.xray(obj)

        # verify
        self.assertEqual([
            mock.call.getScript(
                'vnd.sun.star.script:XrayTool._Main.Xray?'
                'language=Basic&location=application'),
            mock.call.getScript().invoke((obj,), tuple(), tuple())
        ], self.oSP.mock_calls)
        self.assertFalse(self.inspector._ignore_xray)

    def test_xray_fail(self):
        # prepare
        self.oSP.getScript.side_effect = ScriptFrameworkErrorException
        obj = mock.Mock()

        # play
        self.inspector.xray(obj)

        # verify
        self.assertEqual([mock.call.getScript(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')],
            self.oSP.mock_calls)
        self.assertTrue(self.inspector._ignore_xray)

    def test_xray_twice(self):
        # prepare
        self.oSP.getScript.side_effect = ScriptFrameworkErrorException
        obj = mock.Mock()

        # play
        self.inspector.use_xray(False)
        self.inspector.xray(obj)

        # verify
        self.assertEqual([mock.call.getScript(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')],
            self.oSP.mock_calls)
        self.assertTrue(self.inspector._ignore_xray)

    @mock.patch("py4lo_helper.create_uno_service")
    def test_mri(self, us):
        # prepare
        obj = mock.Mock()
        oMRI = mock.Mock()
        us.side_effect = [oMRI]

        # play
        self.inspector.mri(obj)

        # verify
        self.assertEqual([
            mock.call.inspect(obj)
        ], oMRI.mock_calls)
        self.assertFalse(self.inspector._ignore_mri)

    @mock.patch("py4lo_helper.create_uno_service")
    def test_mri_twice(self, us):
        # prepare
        obj = mock.Mock()
        oMRI = mock.Mock()
        us.side_effect = [oMRI]

        # play
        self.inspector.mri(obj)
        self.inspector.mri(obj)

        # verify
        self.assertEqual([
            mock.call.inspect(obj),
            mock.call.inspect(obj)
        ], oMRI.mock_calls)
        self.assertFalse(self.inspector._ignore_mri)

    @mock.patch("py4lo_helper.create_uno_service")
    def test_mri_fail(self, us):
        # prepare
        obj = mock.Mock()
        us.side_effect = ScriptFrameworkErrorException

        # play
        self.inspector.mri(obj)

        # verify
        self.assertTrue(self.inspector._ignore_mri)

    @mock.patch("py4lo_helper.create_uno_service")
    def test_mri_fail_twice(self, us):
        # prepare
        obj = mock.Mock()
        us.side_effect = ScriptFrameworkErrorException

        # play
        self.inspector.mri(obj)
        self.inspector.mri(obj)

        # verify
        self.assertTrue(self.inspector._ignore_mri)

    @mock.patch("py4lo_helper.create_uno_service")
    def test_mri_fail_err(self, us):
        # prepare
        obj = mock.Mock()
        us.side_effect = ScriptFrameworkErrorException

        # play
        with self.assertRaises(UnoRuntimeException):
            self.inspector.mri(obj, True)

        # verify
        self.assertTrue(self.inspector._ignore_mri)

    def test_xray2(self):
        self.inspector.use_xray()
        self.oSP.getScript.assert_called_once_with(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')

    def test_xray3(self):
        self.inspector.xray(1)
        self.inspector.xray(2)
        self.oSP.getScript.assert_called_once_with(
            'vnd.sun.star.script:XrayTool._Main.Xray?'
            'language=Basic&location=application')
        self.oSP.getScript.return_value.invoke.assert_has_calls(
            [mock.call((1,), (), ()), mock.call((2,), (), ())])

    def test_html_converter(self):
        self.maxDiff = None

        # PAR 1
        oPar1Enum = mock.Mock(name="oPar1Enum")
        oPar1Enum.hasMoreElements.side_effect = [True, False]
        oPar1Enum.nextElement.side_effect = [
            self._create_chunk("A", CharFontName="Liberation Avec")]
        oPar1 = mock.Mock(name="oPar1", spec=['createEnumeration'])
        oPar1.Types = [mock.Mock(typeName="com.sun.star.container.XEnumerationAccess")]
        oPar1.createEnumeration.side_effect = [oPar1Enum]

        # PAR 2
        oPar2Enum = mock.Mock()
        oPar2Enum.hasMoreElements.side_effect = [True, True, True, True, True,
                                                 False]
        oPar2Enum.nextElement.side_effect = [
            self._create_chunk("B", CharPosture=FontSlant.ITALIC),
            self._create_chunk("C", CharWeight=150, CharEscapementHeight=58,
                               CharEscapement=-15),
            self._create_chunk("C", CharWeight=150, CharEscapementHeight=58,
                               CharEscapement=15),  # script="sup"),
            self._create_chunk("D", CharColor=0xFF0000),
            self._create_chunk("EFG", CharWeight=150)
        ]
        oPar2 = mock.Mock(name="oPar2", spec=['createEnumeration'])
        oPar2.Types = [mock.Mock(typeName="com.sun.star.container.XEnumerationAccess")]
        oPar2.createEnumeration.side_effect = [oPar2Enum]

        oParEnum = mock.Mock(name="oParEnum")
        oParEnum.hasMoreElements.side_effect = [True, True, False]
        oParEnum.nextElement.side_effect = [oPar1, oPar2]
        oTextRange = mock.Mock(name="oTextRange", spec=['createEnumeration'])
        oTextRange.Types = [mock.Mock(typeName="com.sun.star.container.XEnumerationAccess")]
        oTextRange.createEnumeration.side_effect = [oParEnum]

        self.assertEqual(
            ("""<span style='font-family: "Liberation Avec"'>A</span>"""
             "<br>"
             "<span style='font-style: italic'>B</span>"
             "<sub style='font-weight: 600'>C</sub>"
             "<sup style='font-weight: 600'>C</sup>"
             "<span style='color: #ff0000'>D</span>"
             "<span style='font-weight: 600'>EFG</span>"),
            py4lo_helper.HTMLConverter().convert(oTextRange))

    def _create_chunk(self, string: str, **kwargs) -> UnoTextRange:
        d = {
            'CharFontName': 'Liberation Sans', 'CharHeight': 10,
            'CharWeight': 100, 'CharBackColor': -1,
            'CharColor': -1, 'CharEscapementHeight': 100, 'CharEscapement': 0,
            'CharOverline': 0, 'CharStrikeout': 0, 'CharUnderline': 0,
            'CharPosture': 0,
            **kwargs
        }
        return mock.Mock(String=string, **d)


if __name__ == '__main__':
    unittest.main()
