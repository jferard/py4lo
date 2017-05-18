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

def init(xsc):
    Py4LO_helper.xsc = xsc

class Py4LO_helper(unohelper.Base):
    def __init__(self):
        self.doc = Py4LO_helper.xsc.getDocument()
        self.ctxt = uno.getComponentContext()

        self.ctrl = self.doc.CurrentController
        self.frame = self.ctrl.Frame
        self.parent_win = self.frame.ContainerWindow
        self.sm = self.ctxt.getServiceManager()
        self.dsp = self.doc.getScriptProvider()

        mspf = self.uno_service_ctxt("com.sun.star.script.provider.MasterScriptProviderFactory")
        self.msp = mspf.createScriptProvider("")

        self.loader = self.uno_service( "com.sun.star.frame.Desktop" )
        self.reflect = self.uno_service( "com.sun.star.reflection.CoreReflection" )
        self.dispatcher = self.uno_service( "com.sun.star.frame.DispatchHelper" )
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

    def make_pvs(self, d):
        l = []
        for k in d:
            l.append(self.make_pv(k, d[k]))
        return tuple(l)

    # from https://forum.openoffice.org/fr/forum/viewtopic.php?f=15&t=47603# (thanks Bernard !)
    def message_box(self, parent_win, msg_text, msg_title, msg_type=MESSAGEBOX, msg_buttons=BUTTONS_OK):
        sv = self.uno_service_ctxt("com.sun.star.awt.Toolkit")
        my_box = sv.createMessageBox(parent_win, msg_type, msg_buttons, msg_title, msg_text)
        return my_box.execute()

    def uno_service_ctxt(self, sname, args=None):
        if args is None:
            return self.sm.createInstanceWithContext(sname, self.ctxt)
        else:
            return self.sm.createInstanceWithArgumentsAndContext(sname, args, self.ctxt)

    def uno_service(self, sname, args=None, ctxt=None):
        if ctxt is None:
            return self.sm.createInstance(sname)
        else:
            if args is None:
                return self.sm.createInstanceWithContext(sname, ctxt)
            else:
                return self.sm.createInstanceWithArgumentsAndContext(sname, args, ctxt)

    def new_doc(self):
        return self.loader.loadComponentFromURL(
                     "private:factory/scalc", "_blank", 0, () )

    def open_in_calc(self, filename):
        self.loader.loadComponentFromURL(
            uno.systemPathToFileUrl(filename), "_blank", 0, ())

    def get_last_used_row(self, oSheet):
        oCell = oSheet.GetCellByPosition(0, 0)
        oCursor = oSheet.createCursorByRange(oCell)
        oCursor.GotoEndOfUsedArea(True)
        aAddress = oCursor.RangeAddress
        return aAddress.EndRow

    def to_iter(self, oXIndexAccess):
        for i in range(0, oXIndexAccess.getCount()):
            yield(oXIndexAccess.getByIndex(i))

    def to_dict(self, oXNameAccess):
        d = {}
        for name in oXNameAccess.getElementNames():
            d[name] = oXNameAccess.getByName(name)
        return d

    def read_options_from_sheet_name(self, sheet_name, l=lambda s: s):
        oSheet = self.doc.Sheets.getByName(sheet_name)
        oCell = oSheet.getCellByPosition(0, 0)
        oCursor = oSheet.createCursorByRange(oCell)
        oCursor.gotoEndOfUsedArea(True)
        aAddress = oCursor.RangeAddress
        return self.read_options(oSheet, aAddress, l)

    def read_options(self, oSheet, aAddress, l=lambda s: s):
        options = {}
        for r in range(aAddress.StartRow, aAddress.EndRow+1):
            k = oSheet.getCellByPosition(aAddress.StartColumn, r).String.strip()
            v = oSheet.getCellByPosition(aAddress.StartColumn+1, r).String.strip()
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
        oValidation.ShowErrorMessage = True                    #Anomalies bloquantes
        oValidation.ShowList = uno.getConstantByName(        #Liste déroulante
                            "com.sun.star.sheet.TableValidationVisibility.UNSORTED")
        formula = ";".join('"'+a+'"' for a in fields) #Les valeurs autorisées, entre guillemets et ;
        oValidation.setFormula1(formula)
        oValidation.IgnoreBlankCells = allow_blank            #Saisie obligatoire ?
        cell.Validation = oValidation                        #Affecter validation

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
