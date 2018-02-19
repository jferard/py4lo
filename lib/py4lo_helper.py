# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2018 J. Férard <https://github.com/jferard>

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
import os
import uno
import unohelper
# py4lo: if $python_version >= 2.6
# py4lo: if $python_version <= 3.0
import io
# py4lo: endif
# py4lo: endif
from com.sun.star.uno import RuntimeException
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
# py4lo: if $python_version >= 3.0
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
# py4lo: else
MESSAGEBOX = 0
# py4lo: endif

from com.sun.star.lang import Locale

def init(xsc):
    Py4LO_helper.instance = Py4LO_helper.create(xsc)

class Py4LO_helper(unohelper.Base):
    def create(xsc):
        doc = xsc.getDocument()
        ctxt = uno.getComponentContext()

        ctrl = doc.CurrentController
        frame = ctrl.Frame
        parent_win = frame.ContainerWindow
        sm = ctxt.getServiceManager()
        dsp = doc.getScriptProvider()

        mspf = sm.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctxt)
        msp = mspf.createScriptProvider("")

        reflect = sm.createInstance( "com.sun.star.reflection.CoreReflection" )
        dispatcher = sm.createInstance( "com.sun.star.frame.DispatchHelper" )
        loader = sm.createInstance( "com.sun.star.frame.Desktop" )
        return Py4LO_helper(doc, ctxt, ctrl, frame, parent_win, sm, dsp, mspf, msp, reflect, dispatcher, loader)

    def __init__(self, doc, ctxt, ctrl, frame, parent_win, sm, dsp, mspf, msp, reflect, dispatcher, loader):
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
        self.__xray_script = None

    def use_xray(self):
        try:
            self.__xray_script = self.msp.getScript("vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application")
        except:
            raise RuntimeException("\nBasic library Xray is not installed", self.ctxt)

    def xray(self, object):
        if self.__xray_script is None:
            self.use_xray()

        self.__xray_script.invoke((object,), (), ())

    def make_pv(self, name, value):
        pv = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
        pv.Name = name
        pv.Value = value
        return pv

    def make_pvs(self, d={}):
        l = []
        for n, v in d.items():
            l.append(self.make_pv(n, v))
        return tuple(l)

    # from https://forum.openoffice.org/fr/forum/viewtopic.php?f=15&t=47603# (thanks Bernard !)
    def message_box(self, parent_win, msg_text, msg_title, msg_type=MESSAGEBOX, msg_buttons=BUTTONS_OK):
        sv = self.uno_service_ctxt("com.sun.star.awt.Toolkit")
        mb = sv.createMessageBox(parent_win, msg_type, msg_buttons, msg_title, msg_text)
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
                return self.sm.createInstanceWithArgumentsAndContext(sname, args, ctxt)

    def open_in_calc(self, filename):
        self.loader.loadComponentFromURL(
            uno.systemPathToFileUrl(filename), "_blank", 0, ())

    def get_last_used_row(self, oSheet):
        return self.get_used_range(oSheet).EndRow

    def get_used_range(self, oSheet):
        oCell = oSheet.getCellByPosition(0, 0)
        oCursor = oSheet.createCursorByRange(oCell)
        oCursor.gotoEndOfUsedArea(True)
        return oCursor.RangeAddress

    def to_iter(self, oXIndexAccess):
        for i in range(0, oXIndexAccess.getCount()):
            yield(oXIndexAccess.getByIndex(i))

    def to_dict(self, oXNameAccess):
        d = {}
        for name in oXNameAccess.getElementNames():
            d[name] = oXNameAccess.getByName(name)
        return d

    def new_doc(self, t="calc" ):
        """Create a blank new doc"""
        return self.doc_builder(t).build()

    def doc_builder(self, t="calc"):
        return DocBuilder(self, t)

    # l is deprecated
    def read_options_from_sheet_name(self, sheet_name, l=lambda s: s):
        oSheet = self.doc.Sheets.getByName(sheet_name)
        aAddress = self.get_used_range(oSheet)
        return self.read_options(oSheet, aAddress, l)

    # l is deprecated
    def read_options(self, oSheet, aAddress, l=lambda s: s):
        options = {}
        width = aAddress.EndColumn - aAddress.StartColumn
        for r in range(aAddress.StartRow, aAddress.EndRow+1):
            k = oSheet.getCellByPosition(aAddress.StartColumn, r).String.strip()
            if width <= 1:
                v = oSheet.getCellByPosition(aAddress.StartColumn+1, r).String.strip()
            else:
                # from aAddress.StartColumn+1 to aAddress.EndColumn
                v = [oSheet.getCellByPosition(aAddress.StartColumn+1+c, r).String.strip() for c in range(width)]
            if k and v:
                options[l(k)] = l(v)
        return options

    def get_named_cells(self, name):
        return self.doc.NamedRanges.getByName(name).ReferredCells

    def get_named_cell(self, name):
        return self.get_named_cells(name).getCellByPosition(0,0)

    def set_validation_list_by_name(self, cell_name, fields, default_string=None, allow_blank=False):
        oCell = self.get_named_cell(cell_name)
        self.set_validation_list_by_cell(self, oCell, fields, default_string, allow_blank)

    def set_validation_list_by_cell(self, oCell, fields, default_string=None, allow_blank=False):
        oValidation = oCell.Validation
        oValidation.Type = uno.getConstantByName(
                            "com.sun.star.sheet.ValidationType.LIST")
        oValidation.ShowErrorMessage = True
        oValidation.ShowList = uno.getConstantByName(
                            "com.sun.star.sheet.TableValidationVisibility.UNSORTED")
        formula = ";".join('"'+a+'"' for a in fields)
        oValidation.setFormula1(formula)
        oValidation.IgnoreBlankCells = allow_blank
        cell.Validation = oValidation

        if default_string is not None:
            cell.String = default_string


    def add_filter(self, oDoc, oSheet, range_name):
        oController = oDoc.CurrentController
        oAll = oSheet.getCellRangeByName(range_name)
        oController.select(oAll)
        oFrame = oController.Frame
        self.dispatcher.executeDispatch(oFrame, ".uno:DataFilterAutoFilter", "", 0, ())

    def clear_conditional_format(self, oSheet, range_name):
        oColoredColumns = oSheet.getCellRangeByName(range_name)
        oConditionalFormat = oColoredColumns.ConditionalFormat
        oConditionalFormat.clear()

    def conditional_format_on_formulas(self, oSheet, range_name, style_by_formula, source_position=(0,0)):
        oColoredColumns = oSheet.getCellRangeByName(range_name)
        oConditionalFormat = oColoredColumns.ConditionalFormat
        oSrc = oColoredColumns.getCellByPosition(*source_position).CellAddress

        for formula, style in style_by_formula.items():
            oConditionalEntry = self.get_formula_conditional_entry(formula, style, oSrc)
            oConditionalFormat.addNew(oConditionalEntry)

        oColoredColumns.ConditionalFormat = oConditionalFormat

    def get_formula_conditional_entry(self, formula, style_name, oSrc):
        return self.get_conditional_entry(formula, "", FORMULA, style_name, oSrc)

    def get_conditional_entry(self, formula1, formula2, operator, style_name, oSrc):
        return (
            self.make_pvs({"Formula1":formula1, "Formula2":formula2, "Operator":operator, "StyleName":style_name, "SourcePosition":oSrc})
        )

    # from Andrew Pitonyak 5.14 www.openoffice.org/documentation/HOW_TO/various_topics/AndrewMacro.odt
    def find_or_create_number_format_style(self, oDoc, format):
        oFormats = oDoc.getNumberFormats()
        oLocale = Locale()
        formatNum = oFormats.queryKey(format, oLocale, True)
        if formatNum == -1:
            formatNum = oFormats.addNew(format, oLocale)

        return formatNum

    def copy_row_at_index(self, oSheet, row, r):
        for c, value in enumerate(row):
            if type(value) == str:
                oSheet.getCellByPosition(c, r).String = value
            else:
                oSheet.getCellByPosition(c, r).Value = float(value)


class DocBuilder():
    def __init__(self, helper, t):
        """Create a blank new doc"""
        self.__helper = helper
        self.__oDoc = self.__helper.loader.loadComponentFromURL(
                     "private:factory/s"+t, "_blank", 0, () )
        self.__oDoc.lockControllers()

    def build(self):
        self.__oDoc.unlockControllers()
        return self.__oDoc

    def sheet_names(self, sheet_names, expand_if_necessary=True, trunc_if_necessary=True):
        oSheets = self.__oDoc.Sheets
        it = iter(sheet_names)
        s = 0

        try:
            # rename
            while s < oSheets.getCount():
                oSheet = oSheets.getByIndex(s)
                oSheet.setName(next(it)) # may raise a StopIteration
                s += 1

            assert s == oSheets.getCount(), "s={} vs oSheets.getCount()={}".format(s, oSheets.getCount())

            if expand_if_necessary:
                # add
                for sheet_name in it:
                    oSheets.insertNewByName(sheet_name, s)
                    s += 1
        except StopIteration: # it
            assert s <= oSheets.getCount(), "s={} vs oSheets.getCount()={}".format(s, oSheets.getCount())
            if trunc_if_necessary:
                self.trunc_to_count(s)

        return self

    def apply_func_to_sheets(self, func):
        oSheets = self.__oDoc.Sheets
        for oSheet in oSheets:
            func(oSheet)
        return self

    def apply_func_list_to_sheets(self, funcs):
        oSheets = self.__oDoc.Sheets
        for func, oSheet in zip(funcs, oSheets):
            func(oSheet)
        return self

    def duplicate_base_sheet(self, func, sheet_names, trunc=True):
        """Create a base sheet and duplicate it n-1 times"""
        oSheets = self.__oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        func(oBaseSheet)
        for s, sheet_name in enumerate(sheet_names, 1):
            oSheets.copyByName(oBaseSheet.Name, sheet_name, s)

        return self

    def make_base_sheet(self, func):
        oSheets = self.__oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        func(oBaseSheet)
        return self

    def duplicate_to(self, n):
        oSheets = self.__oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        for s in range(n+1):
            oSheets.copyByName(oBaseSheet.Name, oBaseSheet.Name+str(s), s)

        return self

    def trunc_to_count(self, final_sheet_count):
        oSheets = self.__oDoc.Sheets
        while final_sheet_count < oSheets.getCount():
            oSheet = oSheets.getByIndex(final_sheet_count)
            oSheets.removeByName(oSheet.getName())

        return self
