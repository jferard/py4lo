# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2020 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   Py4LO is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   THIS FILE IS SUBJECT TO THE "CLASSPATH" EXCEPTION.

   Py4LO is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""

"""py4lo_helper deals with LO objects."""

import os
import uno

try:
    import unohelper
except ImportError:
    import unotools.unohelper

# py4lo: if $python_version >= 2.6
# py4lo: if $python_version <= 3.0
import io
# py4lo: endif
# py4lo: endif
from py4lo_commons import float_to_date
from com.sun.star.uno import RuntimeException
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.sheet.ConditionOperator import FORMULA
from com.sun.star.util import NumberFormat
# py4lo: if $python_version >= 3.0
from com.sun.star.awt.MessageBoxType import MESSAGEBOX

# py4lo: else
MESSAGEBOX = 0
# py4lo: endif

from com.sun.star.lang import Locale


def init(xsc):
    Py4LO_helper.instance = Py4LO_helper.create(xsc)


class Py4LO_helper(unohelper.Base):
    @staticmethod
    def create(xsc):
        doc = xsc.getDocument()
        ctxt = uno.getComponentContext()

        ctrl = doc.CurrentController
        frame = ctrl.Frame
        parent_win = frame.ContainerWindow
        sm = ctxt.getServiceManager()
        dsp = doc.getScriptProvider()

        mspf = sm.createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", ctxt)
        msp = mspf.createScriptProvider("")

        reflect = sm.createInstance("com.sun.star.reflection.CoreReflection")
        dispatcher = sm.createInstance("com.sun.star.frame.DispatchHelper")
        loader = sm.createInstance("com.sun.star.frame.Desktop")
        return Py4LO_helper(doc, ctxt, ctrl, frame, parent_win, sm, dsp, mspf,
                            msp, reflect, dispatcher, loader)

    def __init__(self, doc, ctxt, ctrl, frame, parent_win, sm, dsp, mspf, msp,
                 reflect, dispatcher, loader):
        self.doc = doc
        self.ctxt = ctxt
        self.ctrl = ctrl
        self.frame = frame
        self.parent_win = parent_win
        self.sm = sm
        self.dsp = dsp
        self.mspf = mspf
        self.msp = msp
        self.reflect = reflect
        self.dispatcher = dispatcher
        self.loader = loader
        self._xray_script = None
        self._ignore_xray = False
        self._oMRI = None
        self._ignore_mri = False

    def use_xray(self, fail_on_error=False):
        """
        Try to load Xray lib.
        :param fail_on_error: Should this function fail on error
        :raises RuntimeException: if Xray is not avaliable and `fail_on_error` is True.
        """
        try:
            self._xray_script = self.msp.getScript(
                "vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application")
        except:
            if fail_on_error:
                raise RuntimeException("\nBasic library Xray is not installed",
                                       self.ctxt)
            else:
                self._ignore_xray = True

    def xray(self, object, fail_on_error=False):
        """
        Xray an object. Loads dynamically the lib if possible.
        :param fail_on_error: Should this function fail on error
        :raises RuntimeException: if Xray is not avaliable and `fail_on_error` is True.
        """
        if self._ignore_xray:
            return

        if self._xray_script is None:
            self.use_xray(fail_on_error)
            if self._ignore_xray:
                return

        self._xray_script.invoke((object,), (), ())

    def mri(self, object, fail_on_error=False):
        """
        MRI an object
        @param object: the object
        """
        if self._ignore_mri:
            return

        if self._oMRI is None:
            try:
                self._oMRI = self.uno_service("mytools.Mri")
            except:
                if fail_on_error:
                    raise RuntimeException("\nMRI is not installed",
                                           self.ctxt)
                else:
                    self._ignore_mri = True
        self._oMRI.inspect(object)

    # from https://forum.openoffice.org/fr/forum/viewtopic.php?f=15&t=47603# (thanks Bernard !)
    def message_box(self, parent_win, msg_text, msg_title, msg_type=MESSAGEBOX,
                    msg_buttons=BUTTONS_OK):
        sv = self.uno_service_ctxt("com.sun.star.awt.Toolkit")
        mb = sv.createMessageBox(parent_win, msg_type, msg_buttons, msg_title,
                                 msg_text)
        return mb.execute()

    def uno_service_ctxt(self, sname, args=None):
        return self.uno_service(sname, args, self.ctxt)

    def uno_service(self, sname, args=None, ctxt=None):
        if ctxt is None:
            return self.sm.createInstance(sname)
        else:
            if args is None:
                return self.sm.createInstanceWithContext(sname, ctxt)
            else:
                return self.sm.createInstanceWithArgumentsAndContext(sname,
                                                                     args, ctxt)

    def open_in_calc(self, filename):
        """
        Open a document in calc
        :param filename: the name of the file to open
        :return: a reference on the doc
        """
        return self.loader.loadComponentFromURL(
            uno.systemPathToFileUrl(filename), "_blank", 0, ())

    def new_doc(self, t="calc"):
        """Create a blank new doc"""
        return self.doc_builder(t).build()

    def doc_builder(self, t="calc"):
        return DocBuilder(self, t)

    # l is deprecated
    def read_options_from_sheet_name(self, sheet_name, l=lambda s: s):
        oSheet = self.doc.Sheets.getByName(sheet_name)
        oRangeAddress = get_used_range_address(oSheet)
        return read_options(oSheet, oRangeAddress, l)

    def get_named_cells(self, name):
        return self.doc.NamedRanges.getByName(name).ReferredCells

    def get_named_cell(self, name):
        return self.get_named_cells(name).getCellByPosition(0, 0)

    def add_filter(self, oDoc, oSheet, range_name):
        oController = oDoc.CurrentController
        oAll = oSheet.getCellRangeByName(range_name)
        oController.select(oAll)
        oFrame = oController.Frame
        self.dispatcher.executeDispatch(oFrame, ".uno:DataFilterAutoFilter", "",
                                        0, ())

    def set_validation_list_by_name(self, cell_name, fields,
                                    default_string=None, allow_blank=False):
        oCell = self.get_named_cell(cell_name)
        set_validation_list_by_cell(oCell, fields, default_string, allow_blank)


class DocBuilder:
    def __init__(self, helper, t):
        """Create a blank new doc"""
        self._helper = helper
        self._oDoc = self._helper.loader.loadComponentFromURL(
            "private:factory/s" + t, "_blank", 0, ())
        self._oDoc.lockControllers()

    def build(self):
        self._oDoc.unlockControllers()
        return self._oDoc

    def sheet_names(self, sheet_names, expand_if_necessary=True,
                    trunc_if_necessary=True):
        oSheets = self._oDoc.Sheets
        it = iter(sheet_names)
        s = 0

        try:
            # rename
            while s < oSheets.getCount():
                oSheet = oSheets.getByIndex(s)
                oSheet.setName(next(it))  # may raise a StopIteration
                s += 1

            assert s == oSheets.getCount(), "s={} vs oSheets.getCount()={}".format(
                s, oSheets.getCount())

            if expand_if_necessary:
                # add
                for sheet_name in it:
                    oSheets.insertNewByName(sheet_name, s)
                    s += 1
        except StopIteration:  # it
            assert s <= oSheets.getCount(), "s={} vs oSheets.getCount()={}".format(
                s, oSheets.getCount())
            if trunc_if_necessary:
                self.trunc_to_count(s)

        return self

    def apply_func_to_sheets(self, func):
        oSheets = self._oDoc.Sheets
        for oSheet in oSheets:
            func(oSheet)
        return self

    def apply_func_list_to_sheets(self, funcs):
        oSheets = self._oDoc.Sheets
        for func, oSheet in zip(funcs, oSheets):
            func(oSheet)
        return self

    def duplicate_base_sheet(self, func, sheet_names, trunc=True):
        """Create a base sheet and duplicate it n-1 times"""
        oSheets = self._oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        func(oBaseSheet)
        for s, sheet_name in enumerate(sheet_names, 1):
            oSheets.copyByName(oBaseSheet.Name, sheet_name, s)

        return self

    def make_base_sheet(self, func):
        oSheets = self._oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        func(oBaseSheet)
        return self

    def duplicate_to(self, n):
        oSheets = self._oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        for s in range(n + 1):
            oSheets.copyByName(oBaseSheet.Name, oBaseSheet.Name + str(s), s)

        return self

    def trunc_to_count(self, final_sheet_count):
        oSheets = self._oDoc.Sheets
        while final_sheet_count < oSheets.getCount():
            oSheet = oSheets.getByIndex(final_sheet_count)
            oSheets.removeByName(oSheet.getName())

        return self


def make_pv(name, value):
    pv = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    pv.Name = name
    pv.Value = value
    return pv


def make_pvs(d={}):
    return tuple(make_pv(n, v) for n, v in d.items())


def get_last_used_row(oSheet):
    return get_used_range_address(oSheet).EndRow


def get_used_range_address(oSheet):
    oCell = oSheet.getCellByPosition(0, 0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    return oCursor.RangeAddress


def get_used_range(oSheet):
    oRangeAddress = get_used_range_address(oSheet)
    return oSheet.getCellRangeByPosition(oRangeAddress.StartColumn,
                                         oRangeAddress.StartRow,
                                         oRangeAddress.EndColumn,
                                         oRangeAddress.EndRow)


def data_array(oSheet):
    return get_used_range(oSheet).DataArray


def to_iter(oXIndexAccess):
    for i in range(0, oXIndexAccess.getCount()):
        yield oXIndexAccess.getByIndex(i)


def to_dict(oXNameAccess):
    d = {}
    for name in oXNameAccess.getElementNames():
        d[name] = oXNameAccess.getByName(name)
    return d


# l is deprecated
def read_options(oSheet, aAddress, l=lambda s: s):
    options = {}
    width = aAddress.EndColumn - aAddress.StartColumn
    for r in range(aAddress.StartRow, aAddress.EndRow + 1):
        k = oSheet.getCellByPosition(aAddress.StartColumn, r).String.strip()
        if width <= 1:
            v = oSheet.getCellByPosition(aAddress.StartColumn + 1,
                                         r).String.strip()
        else:
            # from aAddress.StartColumn+1 to aAddress.EndColumn
            v = [oSheet.getCellByPosition(aAddress.StartColumn + 1 + c,
                                          r).String.strip() for c in
                 range(width)]
        if k and v:
            options[l(k)] = l(v)
    return options


def set_validation_list_by_cell(oCell, fields, default_string=None,
                                allow_blank=False):
    oValidation = oCell.Validation
    oValidation.Type = uno.getConstantByName(
        "com.sun.star.sheet.ValidationType.LIST")
    oValidation.ShowErrorMessage = True
    oValidation.ShowList = uno.getConstantByName(
        "com.sun.star.sheet.TableValidationVisibility.UNSORTED")
    formula = ";".join('"' + a + '"' for a in fields)
    oValidation.setFormula1(formula)
    oValidation.IgnoreBlankCells = allow_blank
    oCell.Validation = oValidation

    if default_string is not None:
        oCell.String = default_string


def clear_conditional_format(oSheet, range_name):
    oColoredColumns = oSheet.getCellRangeByName(range_name)
    oConditionalFormat = oColoredColumns.ConditionalFormat
    oConditionalFormat.clear()


def conditional_format_on_formulas(oSheet, range_name, style_by_formula,
                                   source_position=(0, 0)):
    oColoredColumns = oSheet.getCellRangeByName(range_name)
    oConditionalFormat = oColoredColumns.ConditionalFormat
    oSrc = oColoredColumns.getCellByPosition(*source_position).CellAddress

    for formula, style in style_by_formula.items():
        oConditionalEntry = get_formula_conditional_entry(formula, style, oSrc)
        oConditionalFormat.addNew(oConditionalEntry)

    oColoredColumns.ConditionalFormat = oConditionalFormat


def get_formula_conditional_entry(formula, style_name, oSrc):
    return get_conditional_entry(formula, "", FORMULA, style_name, oSrc)


def get_conditional_entry(formula1, formula2, operator, style_name, oSrc):
    return (
        make_pvs(
            {"Formula1": formula1, "Formula2": formula2, "Operator": operator,
             "StyleName": style_name,
             "SourcePosition": oSrc})
    )


# from Andrew Pitonyak 5.14 www.openoffice.org/documentation/HOW_TO/various_topics/AndrewMacro.odt
def find_or_create_number_format_style(oDoc, format):
    oFormats = oDoc.getNumberFormats()
    oLocale = Locale()
    formatNum = oFormats.queryKey(format, oLocale, True)
    if formatNum == -1:
        formatNum = oFormats.addNew(format, oLocale)

    return formatNum


def copy_row_at_index(oSheet, row, r):
    for c, value in enumerate(row):
        if type(value) == str:
            oSheet.getCellByPosition(c, r).String = value
        else:
            oSheet.getCellByPosition(c, r).Value = float(value)


def get_range_size(oRange):
    """
    Useful for `oRange.getRangeCellByPosition(...)`.

    :param oRange: a SheetCellRange object
    :return: width and height of the range.
    """
    oAddress = oRange.RangeAddress
    width = oAddress.EndColumn - oAddress.StartColumn + 1
    height = oAddress.EndRow - oAddress.StartRow + 1
    return width, height


def get_doc(oCell):
    """
    Find the document that owns this cell.

    @param oCell: a cell
    @return: the document
    """
    return oCell.Spreadsheet.DrawPage.Forms.Parent


def type_cell(oCell, oFormats=None):
    """
    Type a cell value
    @param oCell: the cell
    @return: the cell value as int or text
    """
    if oFormats is None:
        oFormats = get_doc(oCell).NumberFormats
    key = oCell.NumberFormat
    cell_type = oCell.getType().value
    if cell_type == 'FORMULA':
        cell_type = oCell.FormulaResultType.value

    if cell_type == 'EMPTY':
        return None
    elif cell_type == 'TEXT':
        return oCell.String
    else:
        cell_data_type = oFormats.getByKey(key).Type
        if cell_data_type in {NumberFormat.DATE, NumberFormat.DATETIME,
                              NumberFormat.TIME}:
            return float_to_date(oCell.Value)
        elif cell_data_type in {NumberFormat.CURRENCY, NumberFormat.FRACTION,
                                NumberFormat.NUMBER, NumberFormat.PERCENT,
                                NumberFormat.SCIENTIFIC}:
            return oCell.Value
        elif cell_data_type == NumberFormat.LOGICAL:
            return bool(oCell.Value)
        else:
            return oCell.String


class reader:
    def __init__(self, oSheet, type_cell=type_cell):
        self._oSheet = oSheet
        self._type_cell = type_cell
        self.line_num = 0
        self._oFormats = oSheet.DrawPage.Forms.Parent.NumberFormats
        self._oRangeAddress = get_used_range_address(oSheet)

    def __iter__(self):
        return self

    def __next__(self):
        i = self._oRangeAddress.StartRow + self.line_num
        if i > self._oRangeAddress.EndRow:
            raise StopIteration

        self.line_num += 1
        if self._type_cell is None:
            row = [self._oSheet.getCellByPosition(j, i).String
                   for j in range(self._oRangeAddress.StartColumn,
                                  self._oRangeAddress.EndColumn + 1)]
        else:
            row = [self._type_cell(self._oSheet.getCellByPosition(j, i), self._oFormats)
                   for j in range(self._oRangeAddress.StartColumn,
                                  self._oRangeAddress.EndColumn + 1)]

        # left strip the row
        i = len(row) - 1
        while row[i] is None and i > 0:
            i -= 1
        return row[:i + 1]


class dict_reader:
    def __init__(self, oSheet, fieldnames=None, restkey=None, restval=None):
        self._reader = reader(oSheet)
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
