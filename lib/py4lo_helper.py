# -*- coding: utf-8 -*-
# Py4LO - Python Toolkit For LibreOffice Calc
#       Copyright (C) 2016-2022 J. Férard <https://github.com/jferard>
#
#    This file is part of Py4LO.
#
#    Py4LO is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    THIS FILE IS SUBJECT TO THE "CLASSPATH" EXCEPTION.
#
#    Py4LO is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""py4lo_helper deals with LO objects."""

from enum import Enum

from typing import (Any, Optional, List, cast, Callable, Mapping, Tuple,
                    Iterator, Union)

from py4lo_typing import (UnoSpreadsheet, UnoController, UnoContext, UnoService,
                          UnoSheet, UnoRangeAddress, UnoRange, UnoCell,
                          UnoObject, DATA_ARRAY, UnoCellAddress,
                          UnoPropertyValue, DATA_ROW, UnoXScriptContext)

import os
import uno

try:
    import unohelper


    class FrameSearchFlag:
        from com.sun.star.frame.FrameSearchFlag import (
            AUTO, PARENT, SELF, CHILDREN, CREATE, SIBLINGS, TASKS, ALL, GLOBAL)
except ImportError:
    import unotools.unohelper

# py4lo: if $python_version >= 2.6
# py4lo: if $python_version <= 3.0
# py4lo: endif
# py4lo: endif
from com.sun.star.uno import RuntimeException as UnoRuntimeException, \
    Exception as UnoException


class ConditionOperator:
    from com.sun.star.sheet.ConditionOperator import (FORMULA, )


from com.sun.star.script.provider import ScriptFrameworkErrorException


class MessageBoxButtons:
    from com.sun.star.awt.MessageBoxButtons import (BUTTONS_OK, )


class MessageBoxType:
    from com.sun.star.awt.MessageBoxType import (MESSAGEBOX, )


from com.sun.star.lang import Locale

provider = cast(Optional["_ObjectProvider"], None)
_inspect = cast(Optional["_Inspector"], None)
xray = cast(Optional[Callable], None)
mri = cast(Optional[Callable], None)


def init(xsc: UnoXScriptContext):
    """
    Mandatory call from entry with XSCRIPTCONTEXT as argument.
    @param xsc: XSCRIPTCONTEXT
    """
    global provider, _inspect, xray, mri
    provider = _ObjectProvider.create(xsc)
    _inspect = _Inspector(provider)
    xray = _inspect.xray
    mri = _inspect.mri


class _ObjectProvider:
    """
    Lazy object provider.
    """

    @staticmethod
    def create(xsc: UnoXScriptContext):
        doc = xsc.getDocument()
        controller = doc.CurrentController
        frame = controller.Frame
        parent_win = frame.ContainerWindow
        script_provider = doc.getScriptProvider()
        ctxt = xsc.getComponentContext()
        service_manager = ctxt.getServiceManager()
        desktop = xsc.getDesktop()
        return _ObjectProvider(doc, controller, frame, parent_win,
                               script_provider, ctxt, service_manager, desktop)

    def __init__(self, doc: UnoSpreadsheet, controller: UnoController, frame,
                 parent_win, script_provider,
                 ctxt: UnoContext, service_manager, desktop):
        self.doc = doc
        self.controller = controller
        self.frame = frame
        self.parent_win = parent_win
        self.script_provider = script_provider
        self.ctxt = ctxt
        self.service_manager = service_manager
        self.desktop = desktop
        self._script_provider_factory = None
        self._script_provider = None
        self._reflect = None
        self._dispatcher = None

    def get_script_provider_factory(self):
        """
        > This service is used to create MasterScriptProviders
        @return:
        """
        if self._script_provider_factory is None:
            self._script_provider_factory = self.service_manager.createInstanceWithContext(
                "com.sun.star.script.provider.MasterScriptProviderFactory",
                self.ctxt)
        return self._script_provider_factory

    def get_script_provider(self):
        """
        > This interface provides a factory for obtaining objects implementing the XScript interface

        @return:
        """
        if self._script_provider is None:
            self._script_provider = self.get_script_provider_factory().createScriptProvider(
                "")
        return self._script_provider

    @property
    def reflect(self):
        """
        > This service is the implementation of the reflection API

        @return:
        """
        if self._reflect is None:
            self._reflect = self.service_manager.createInstance(
                "com.sun.star.reflection.CoreReflection")
        return self._reflect

    @property
    def dispatcher(self):
        """
        > provides an easy way to dispatch a URL using one call instead of multiple ones.
        @return:
        """
        if self._dispatcher is None:
            self._dispatcher = self.service_manager.createInstance(
                "com.sun.star.frame.DispatchHelper")
        return self._dispatcher


class _Inspector:
    def __init__(self, provider: _ObjectProvider):
        self._provider = provider
        self._xray_script = None
        self._ignore_xray = False
        self._oMRI = None
        self._ignore_mri = False

    def use_xray(self, fail_on_error: bool = False):
        """
        Try to load Xray lib.
        :param fail_on_error: Should this function fail on error
        :raises UnoRuntimeException: if Xray is not avaliable and `fail_on_error` is True.
        """
        try:
            self._xray_script = self._provider.script_provider.getScript(
                "vnd.sun.star.script:XrayTool._Main.Xray?"
                "language=Basic&location=application")
        except ScriptFrameworkErrorException:
            if fail_on_error:
                raise UnoRuntimeException(
                    "\nBasic library Xray is not installed",
                    self._provider.ctxt)
            else:
                self._ignore_xray = True

    def xray(self, obj: Any, fail_on_error: bool = False):
        """
        Xray an obj. Loads dynamically the lib if possible.
        :@param obj: the obj
        :@param fail_on_error: Should this function fail on error
        :@raises RuntimeException: if Xray is not avaliable and `fail_on_error` is True.
        """
        if self._ignore_xray:
            return

        if self._xray_script is None:
            self.use_xray(fail_on_error)
            if self._ignore_xray:
                return

        self._xray_script.invoke((obj,), (), ())

    def mri(self, obj: Any, fail_on_error: bool = False):
        """
        MRI an object
        @param fail_on_error:
        @param obj: the object
        """
        if self._ignore_mri:
            return

        if self._oMRI is None:
            try:
                self._oMRI = uno_service("mytools.Mri")
            except UnoException:
                if fail_on_error:
                    raise UnoRuntimeException("\nMRI is not installed",
                                              self._provider.ctxt)
                else:
                    self._ignore_mri = True
        self._oMRI.inspect(obj)


def uno_service(sname: str, args: Optional[List[Any]] = None,
                ctxt: Optional[UnoContext] = None) -> UnoService:
    sm = provider.service_manager
    if ctxt is None:
        return sm.createInstance(sname)
    else:
        if args is None:
            return sm.createInstanceWithContext(sname, ctxt)
        else:
            return sm.createInstanceWithArgumentsAndContext(sname, args, ctxt)


def uno_service_ctxt(sname: str,
                     args: Optional[List[Any]] = None) -> UnoService:
    return uno_service(sname, args, provider.ctxt)


def message_box(msg_text: str, msg_title: str,
                msg_type=MessageBoxType.MESSAGEBOX,
                msg_buttons=MessageBoxButtons.BUTTONS_OK, parent_win=None):
    # from https://forum.openoffice.org/fr/forum/viewtopic.php?f=15&t=47603#
    # (thanks Bernard !)
    toolkit = uno_service_ctxt("com.sun.star.awt.Toolkit")
    if parent_win is None:
        parent_win = provider.parent_win
    mb = toolkit.createMessageBox(parent_win, msg_type, msg_buttons, msg_title,
                                  msg_text)
    return mb.execute()


###
# Open a document
###
NEW_CALC_DOCUMENT = "private:factory/scalc"
NEW_WRITER_DOCUMENT = "private:factory/swriter"


# special targets
class Target(str, Enum):
    BLANK = "_blank"  # always creates a new frame
    DEFAULT = "_default"  # special UI functionality (e.g. detecting of already loaded documents, using of empty frames of creating of new top frames as fallback)
    SELF = "_self"  # means frame himself
    PARENT = "_parent"  # address direct parent of frame
    TOP = "_top"  # indicates top frame of current path in tree
    BEAMER = "_beamer"  # means special sub frame


def open_in_calc(filename: str, target: str = Target.BLANK,
                 frame_flags=FrameSearchFlag.AUTO,
                 **kwargs) -> UnoSpreadsheet:
    """
    Open a document in calc
    :param filename: the name of the file to open
    :param target: "the name of the frame to view the document in" or a special target
    :param frame_flags: where to search the frame
    :param kwargs: les paramètres d'ouverture
    :return: a reference on the doc
    """
    url = uno.systemPathToFileUrl(os.path.realpath(filename))
    if kwargs:
        params = make_pvs(kwargs)
    else:
        params = ()
    return provider.desktop.loadComponentFromURL(url, target, frame_flags,
                                                 params)


# Create a document
###
def doc_builder(t: str = "calc") -> "DocBuilder":
    return DocBuilder(t)


def new_doc(t: str = "calc") -> UnoSpreadsheet:
    """Create a blank new doc"""
    return doc_builder(t).build()


def add_filter(oDoc: UnoSpreadsheet, oSheet: UnoSheet, range_name: str):
    oController = oDoc.CurrentController
    oAll = oSheet.getCellRangeByName(range_name)
    oController.select(oAll)
    oFrame = oController.Frame
    provider.dispatcher.executeDispatch(
        oFrame, ".uno:DataFilterAutoFilter", "", 0, tuple())


def read_options_from_sheet_name(
        oDoc: UnoSpreadsheet, sheet_name: str,
        apply: Callable[[Tuple[str, str]], Tuple[str, str]] = lambda k_v: k_v):
    oSheet = oDoc.Sheets.getByName(sheet_name)
    oRangeAddress = get_used_range_address(oSheet)
    return read_options(oSheet, oRangeAddress, apply)


def get_named_cells(oDoc: UnoSpreadsheet, name: str) -> UnoRange:
    return oDoc.NamedRanges.getByName(name).ReferredCells


def get_named_cell(oDoc: UnoSpreadsheet, name: str) -> UnoCell:
    return get_named_cells(oDoc, name).getCellByPosition(0, 0)


def set_validation_list_by_name(
        oDoc: UnoSpreadsheet, cell_name: str, fields: List[str],
        default_string: Optional[str] = None, allow_blank: bool = False):
    oCell = get_named_cell(oDoc, cell_name)
    set_validation_list_by_cell(oCell, fields, default_string, allow_blank)


class DocBuilder:
    def __init__(self, t: str):
        """Create a blank new doc"""
        self._oDoc = provider.desktop.loadComponentFromURL(
            "private:factory/s" + t, "_blank", 0, ())
        self._oDoc.lockControllers()

    def build(self) -> UnoSpreadsheet:
        self._oDoc.unlockControllers()
        return self._oDoc

    def sheet_names(self, sheet_names: str, expand_if_necessary: bool = True,
                    trunc_if_necessary: bool = True) -> "DocBuilder":
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

    def apply_func_to_sheets(
            self, func: Callable[[UnoSheet], None]) -> "DocBuilder":
        oSheets = self._oDoc.Sheets
        for oSheet in oSheets:
            func(oSheet)
        return self

    def apply_func_list_to_sheets(
            self, funcs: List[Callable[[UnoSheet], None]]) -> "DocBuilder":
        oSheets = self._oDoc.Sheets
        for func, oSheet in zip(funcs, oSheets):
            func(oSheet)
        return self

    def duplicate_base_sheet(self, func: Callable[[UnoSheet], None],
                             sheet_names: List[str], trunc: bool = True):
        """Create a base sheet and duplicate it n-1 times"""
        oSheets = self._oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        func(oBaseSheet)
        for s, sheet_name in enumerate(sheet_names, 1):
            oSheets.copyByName(oBaseSheet.Name, sheet_name, s)

        return self

    def make_base_sheet(self, func: Callable[[UnoSheet], None]) -> "DocBuilder":
        oSheets = self._oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        func(oBaseSheet)
        return self

    def duplicate_to(self, n: int) -> "DocBuilder":
        oSheets = self._oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        for s in range(n + 1):
            oSheets.copyByName(oBaseSheet.Name, oBaseSheet.Name + str(s), s)

        return self

    def trunc_to_count(self, final_sheet_count: int) -> "DocBuilder":
        oSheets = self._oDoc.Sheets
        while final_sheet_count < oSheets.getCount():
            oSheet = oSheets.getByIndex(final_sheet_count)
            oSheets.removeByName(oSheet.getName())

        return self


def make_pv(name: str, value: str) -> UnoPropertyValue:
    pv = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    pv.Name = name
    pv.Value = value
    return pv


def make_pvs(d: Optional[Mapping[str, str]] = None) -> Tuple[
    UnoPropertyValue, ...]:
    if d is None:
        return tuple()
    else:
        return tuple(make_pv(n, v) for n, v in d.items())


def get_last_used_row(oSheet: UnoSheet) -> int:
    return get_used_range_address(oSheet).EndRow


def get_used_range_address(oSheet: UnoSheet) -> UnoRangeAddress:
    oCell = oSheet.getCellByPosition(0, 0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    return oCursor.RangeAddress


def get_used_range(oSheet: UnoSheet) -> UnoRange:
    oRangeAddress = get_used_range_address(oSheet)
    return narrow_range_to_address(oSheet, oRangeAddress)


def narrow_range_to_address(
        oSheet: UnoSheet, oRangeAddress: UnoRangeAddress) -> UnoRange:
    """
    Useful to copy a data array from one sheet to another
    @param oSheet:
    @param oRangeAddress:
    @return:
    """
    return oSheet.getCellRangeByPosition(
        oRangeAddress.StartColumn, oRangeAddress.StartRow,
        oRangeAddress.EndColumn, oRangeAddress.EndRow)


def data_array(oSheet: UnoSheet) -> DATA_ARRAY:
    return get_used_range(oSheet).DataArray


def to_iter(oXIndexAccess: UnoObject) -> Iterator[UnoObject]:
    for i in range(0, oXIndexAccess.getCount()):
        yield oXIndexAccess.getByIndex(i)


def to_dict(oXNameAccess: UnoObject) -> Mapping[str, UnoObject]:
    d = {}
    for name in oXNameAccess.getElementNames():
        d[name] = oXNameAccess.getByName(name)
    return d


def read_options(oSheet: UnoSpreadsheet, aAddress: UnoRangeAddress,
                 apply: Callable[[Tuple[str, str]],
                                 Tuple[str, str]] = lambda k_v: k_v
                 ) -> Mapping[str, str]:
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
            k, v = apply((k, v))
            options[k] = v
    return options


def set_validation_list_by_cell(
        oCell: UnoCell, fields: List[str], default_string: Optional[str] = None,
        allow_blank: bool = False):
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


def clear_conditional_format(oSheet: UnoSheet, range_name: str):
    oColoredColumns = oSheet.getCellRangeByName(range_name)
    oConditionalFormat = oColoredColumns.ConditionalFormat
    oConditionalFormat.clear()


def conditional_format_on_formulas(
        oSheet: UnoSheet, range_name: str, style_by_formula,
        source_position: Tuple[int, int] = (0, 0)):
    oColoredColumns = oSheet.getCellRangeByName(range_name)
    oConditionalFormat = oColoredColumns.ConditionalFormat
    oSrcAddress = oColoredColumns.getCellByPosition(
        *source_position).CellAddress

    for formula, style in style_by_formula.items():
        oConditionalEntry = get_formula_conditional_entry(formula, style,
                                                          oSrcAddress)
        oConditionalFormat.addNew(oConditionalEntry)

    oColoredColumns.ConditionalFormat = oConditionalFormat


def get_formula_conditional_entry(
        formula: str, style_name: str, oSrcAddress: UnoCellAddress
) -> Tuple[UnoPropertyValue, ...]:
    return get_conditional_entry(formula, "", ConditionOperator.FORMULA,
                                 style_name, oSrcAddress)


def get_conditional_entry(
        formula1: str, formula2: str, operator: str, style_name: str,
        oSrcAddress: UnoCellAddress) -> Tuple[UnoPropertyValue, ...]:
    return make_pvs(
        {"Formula1": formula1, "Formula2": formula2, "Operator": operator,
         "StyleName": style_name,
         "SourcePosition": oSrcAddress}
    )


# from Andrew Pitonyak 5.14 www.openoffice.org/documentation/HOW_TO/various_topics/AndrewMacro.odt
def find_or_create_number_format_style(oDoc: UnoSpreadsheet, format: str
                                       ) -> int:
    oFormats = oDoc.getNumberFormats()
    oLocale = Locale()
    formatNum = oFormats.queryKey(format, oLocale, True)
    if formatNum == -1:
        formatNum = oFormats.addNew(format, oLocale)

    return formatNum


def copy_row_at_index(oSheet: UnoSheet, row: DATA_ROW, r: int):
    oRange = oSheet.getCellRangeByPosition(0, 0, len(row) - 1, 0)
    oRange.DataArray = row


def get_range_size(oRange: UnoRange) -> Tuple[int, int]:
    """
    Useful for `oRange.getRangeCellByPosition(...)`.

    :param oRange: a SheetCellRange obj
    :return: width and height of the range.
    """
    oAddress = oRange.RangeAddress
    width = oAddress.EndColumn - oAddress.StartColumn + 1
    height = oAddress.EndRow - oAddress.StartRow + 1
    return width, height


def get_doc(oRange: Union[UnoCell, UnoRange]) -> UnoSpreadsheet:
    """
    Find the document that owns this cell.

    @param oRange: a cell or a range
    @return: the document
    """
    return parent_doc(oRange.Spreadsheet)


def get_cell_type(oCell: str) -> str:
    """
    @param oCell: the cell
    @return: 'EMPTY', 'TEXT', 'VALUE'
    """
    cell_type = oCell.getType().value
    if cell_type == 'FORMULA':
        cell_type = oCell.FormulaResultType.value

    return cell_type


def parent_doc(oSheet: UnoSheet) -> UnoSpreadsheet:
    """
    @param oSheet: the sheet
    @return: the document to which this sheet belongs
    """
    return oSheet.DrawPage.Forms.Parent


def create_filter(oRange: UnoRange):
    """
    Create a new filter
    @param oRange: the range to filter
    """
    oDoc = get_doc(oRange)
    oDoc.CurrentController.select(oRange)
    oDispatchHelper = uno_service("com.sun.star.frame.DispatchHelper")
    oDispatchHelper.executeDispatch(oDoc.CurrentController.Frame,
                                    ".uno:DataFilterAutoFilter", "", 0, [])
