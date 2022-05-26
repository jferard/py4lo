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

import os
from enum import Enum
from pathlib import Path
from typing import (Any, Optional, List, cast, Callable, Mapping, Tuple,
                    Iterator, Union)

from py4lo_typing import (UnoSpreadsheet, UnoController, UnoContext, UnoService,
                          UnoSheet, UnoRangeAddress, UnoRange, UnoCell,
                          UnoObject, DATA_ARRAY, UnoCellAddress,
                          UnoPropertyValue, DATA_ROW, UnoXScriptContext,
                          UnoColumn, UnoStruct, UnoEnum, UnoRow)

try:
    import uno


    class FrameSearchFlag:
        from com.sun.star.frame.FrameSearchFlag import (
            AUTO, PARENT, SELF, CHILDREN, CREATE, SIBLINGS, TASKS, ALL, GLOBAL)


    class ConditionOperator:
        from com.sun.star.sheet.ConditionOperator import (FORMULA, )


    class FontWeight:
        from com.sun.star.awt.FontWeight import (BOLD, )


    class BorderLineStyle:
        from com.sun.star.table.BorderLineStyle import (SOLID, )


    class ValidationType:
        from com.sun.star.sheet.ValidationType import (LIST, )


    class TableValidationVisibility:
        from com.sun.star.sheet.TableValidationVisibility import (
            SORTEDASCENDING, UNSORTED)


    from com.sun.star.script.provider import ScriptFrameworkErrorException
    from com.sun.star.uno import (RuntimeException as UnoRuntimeException,
                                  Exception as UnoException)

except ImportError:
    class FrameSearchFlag:
        AUTO = None


    class BorderLineStyle:
        SOLID = None

###############################################################################
# BASE
###############################################################################
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
            self._script_provider_factory = \
                self.service_manager.createInstanceWithContext(
                    "com.sun.star.script.provider.MasterScriptProviderFactory",
                    self.ctxt)
        return self._script_provider_factory

    def get_script_provider(self):
        """
        > This interface provides a factory for obtaining objects implementing
        > the XScript interface

        @return:
        """
        if self._script_provider is None:
            self._script_provider = \
                self.get_script_provider_factory().createScriptProvider("")
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
        > provides an easy way to dispatch a URL using one call instead of
        > multiple ones.
        @return:
        """
        if self._dispatcher is None:
            self._dispatcher = self.service_manager.createInstance(
                "com.sun.star.frame.DispatchHelper")
        return self._dispatcher


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


def to_iter(oXIndexAccess: UnoObject) -> Iterator[UnoObject]:
    for i in range(0, oXIndexAccess.getCount()):
        yield oXIndexAccess.getByIndex(i)


def to_dict(oXNameAccess: UnoObject) -> Mapping[str, UnoObject]:
    d = {}
    for name in oXNameAccess.getElementNames():
        d[name] = oXNameAccess.getByName(name)
    return d


def uno_url_to_path(url: str) -> Optional[Path]:
    """
    Wrapper
    @param url: the url
    @return: the path or None if the url is empty
    """
    if url.strip():
        return Path(uno.fileUrlToSystemPath(url))
    else:
        return None


def uno_path_to_url(path: Union[str, Path]) -> str:
    """
    Wrapper
    @param path: the path
    @return: the url
    """
    return uno.systemPathToFileUrl(str(path))


def parent_doc(oRange: UnoRange) -> UnoSpreadsheet:
    """
    Find the document that owns this range.

    @param oRange: the range (range, sheet, cell)
    @return: the document to which this range belongs
    """
    oSheet = oRange.Speadsheet
    return oSheet.DrawPage.Forms.Parent


def get_cell_type(oCell: UnoCell) -> str:
    """
    @param oCell: the cell
    @return: 'EMPTY', 'TEXT', 'VALUE'
    """
    cell_type = oCell.Type.value
    if cell_type == 'FORMULA':
        cell_type = oCell.FormulaResultType.value

    return cell_type


def get_named_cells(oDoc: UnoSpreadsheet, name: str) -> UnoRange:
    return oDoc.NamedRanges.getByName(name).ReferredCells


def get_named_cell(oDoc: UnoSpreadsheet, name: str) -> UnoCell:
    return get_named_cells(oDoc, name).getCellByPosition(0, 0)


###############################################################################
# OPEN A DOCUMENT
###############################################################################

class NewDocumentUrl(str, Enum):
    Calc = "private:factory/scalc"
    Writer = "private:factory/swriter"
    Draw = "private:factory/sdraw"
    Impress = "private:factory/simpress"
    Math = "private:factory/smath"


# special targets
class Target(str, Enum):
    BLANK = "_blank"  # always creates a new frame
    # special UI functionality (e.g. detecting of already loaded documents,
    # using of empty frames of creating of new top frames as fallback)
    DEFAULT = "_default"
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
    :param target: "the name of the frame to view the document in" or a special
    target
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
####
def doc_builder(
        url: NewDocumentUrl = NewDocumentUrl.Calc,
        taget_frame_name: Target = Target.BLANK,
        search_flags: FrameSearchFlag = FrameSearchFlag.AUTO,
        pvs: List[UnoPropertyValue] = None
) -> "DocBuilder":
    if pvs is None:
        pvs = tuple()
    return DocBuilder(url, taget_frame_name, search_flags, pvs)


def new_doc(url: NewDocumentUrl = NewDocumentUrl.Calc,
            taget_frame_name: Target = Target.BLANK,
            search_flags: FrameSearchFlag = FrameSearchFlag.AUTO,
            pvs: List[UnoPropertyValue] = None) -> UnoSpreadsheet:
    """Create a blank new doc"""
    return doc_builder(url, taget_frame_name, search_flags, pvs).build()


class DocBuilder:
    def __init__(self, url: NewDocumentUrl, taget_frame_name: Target,
                 search_flags: FrameSearchFlag, pvs: List[UnoPropertyValue]):
        """Create a blank new doc"""
        self._oDoc = provider.desktop.loadComponentFromURL(
            url, taget_frame_name, search_flags, pvs)
        self._oDoc.lockControllers()

    def build(self) -> UnoSpreadsheet:
        self._oDoc.unlockControllers()
        return self._oDoc

    def sheet_names(self, sheet_names: List[str],
                    expand_if_necessary: bool = True,
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

            if s != oSheets.getCount():
                raise AssertionError("s={} vs oSheets.getCount()={}".format(
                    s, oSheets.getCount()))

            if expand_if_necessary:
                # add
                for sheet_name in it:
                    oSheets.insertNewByName(sheet_name, s)
                    s += 1
        except StopIteration:  # it
            if s > oSheets.getCount():
                raise AssertionError("s={} vs oSheets.getCount()={}".format(
                    s, oSheets.getCount()))
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


##############################################################################
# STRUCTS
##############################################################################

def make_struct(struct_id: str, **kwargs):
    struct = uno.createUnoStruct(struct_id)
    for k, v in kwargs.items():
        struct.__setattr__(k, v)
    return struct


def make_pv(name: str, value: str) -> UnoPropertyValue:
    """
    @param name: the name of the PropertyValue
    @param value: the value of the PropertyValue
    @return: the PropertyValue
    """
    pv = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    pv.Name = name
    pv.Value = value
    return pv


def make_full_pv(name: str, value: str, handle: int = -1,
                 state: Optional[UnoEnum] = None) -> UnoPropertyValue:
    """
    @param name: the name of the PropertyValue
    @param value: the value of the PropertyValue
    @param handle: the handle
    @param state: the state
    @return: the PropertyValue
    """
    pv = make_pv(name, value)
    if handle != -1:
        pv.Handle = handle
    if state is not None:
        pv.State = state
    return pv


def make_pvs(d: Optional[Mapping[str, str]] = None
             ) -> Tuple[UnoPropertyValue, ...]:
    if d is None:
        return tuple()
    else:
        return tuple(make_pv(n, v) for n, v in d.items())


def make_locale(country: str = "", language: str = "",
                variant: str = "") -> UnoStruct:
    """
    Create a locale

    @param country: ISO 3166 Country Code.
    @param language: ISO 639 Language Code.
    @param variant: BCP 47
    @return: the locale
    """
    locale = uno.createUnoStruct('com.sun.star.lang.Locale')
    locale.Country = country
    if language:
        if variant:
            raise ValueError("Language or Variant")
        locale.Language = language
    elif variant:
        locale.Language = "qlt"
        locale.Variant = variant
    return locale


def make_border(color: int, width: int,
                style: BorderLineStyle = BorderLineStyle.SOLID):
    """
    Create a border
    @param color: the color
    @param width: the width
    @param style: the style
    @return: the border
    """
    border = uno.createUnoStruct("com.sun.star.table.BorderLine2")
    border.Color = color
    border.LineWidth = width
    border.LineStyle = style
    return border


##############################################################################
# RANGES
##############################################################################

def get_last_used_row(oSheet: UnoSheet) -> int:
    return get_used_range_address(oSheet).EndRow


def get_used_range_address(oSheet: UnoSheet) -> UnoRangeAddress:
    """
    @param oSheet: the sheet
    @return: the used range address
    """
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(True)
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


def narrow_range(oRange: UnoRange, narrow_data: bool = False
                 ) -> Optional[UnoRange]:
    """
    Narrow the range to the used range
    @param oRange: the range, usually a row or a column
    @param narrow_data: if True, remove top/bottom blank lines and left/right
     blank colmuns
    @return the narrowed range or None
    """
    oSheet = oRange.Spreadsheet
    oSheetRangeAddress = get_used_range_address(oSheet)
    oRangeAddress = oRange.RangeAddress
    start_column = max(oRangeAddress.StartColumn,
                       oSheetRangeAddress.StartColumn)
    end_column = min(oRangeAddress.EndColumn, oSheetRangeAddress.EndColumn)
    if start_column > end_column:
        return None
    start_row = max(oRangeAddress.StartRow, oSheetRangeAddress.StartRow)
    end_row = min(oRangeAddress.EndRow, oSheetRangeAddress.EndRow)
    if start_row > end_row:
        return None

    oNarrowedRange = oSheet.getCellRangeByPosition(
        start_column, start_row, end_column, end_row)

    if narrow_data:
        data_array = oNarrowedRange.DataArray
        start_row += top_void_row_count(data_array)
        if start_row > end_row:
            return None
        end_row -= bottom_void_row_count(data_array)
        start_column += left_void_row_count(data_array)
        end_column -= right_void_row_count(data_array)
        oNarrowedRange = oSheet.getCellRangeByPosition(
            start_column, start_row, end_column, end_row)

    return oNarrowedRange


##############################################################################
# DATA ARRAY
##############################################################################
def data_array(oSheet: UnoSheet) -> DATA_ARRAY:
    return get_used_range(oSheet).DataArray


def top_void_row_count(data_array: DATA_ARRAY) -> int:
    """
    @param data_array: a data array
    @return: the number of void row at the top
    """
    r0 = 0
    row_count = len(data_array)
    while r0 < row_count and all(v.strip() == "" for v in data_array[r0]):
        r0 += 1
    return r0


def bottom_void_row_count(data_array: DATA_ARRAY) -> int:
    """
    @param data_array: a data array
    @return: the number of void row at the top
    """
    row_count = len(data_array)
    r1 = 0
    # r1 < row_count => row_count - r1 > 0 => row_count - r1 - 1 >= 0
    while r1 < row_count and all(
            v.strip() == "" for v in data_array[row_count - r1 - 1]):
        r1 += 1
    return r1


def left_void_row_count(data_array: DATA_ARRAY) -> int:
    """
    @param data_array: a data array
    @return: the number of void row at the top
    """
    row_count = len(data_array)
    if row_count == 0:
        return 0

    c0 = len(data_array[0])
    for row in data_array:
        c = 0
        while c < c0 and row[c].strip() == "":
            c += 1
        if c < c0:
            c0 = c
    return c0


def right_void_row_count(data_array: DATA_ARRAY) -> int:
    """
    @param data_array: a data array
    @return: the number of void row at the top
    """
    row_count = len(data_array)
    if row_count == 0:
        return 0

    width = len(data_array[0])
    c1 = width
    for row in data_array:
        c = 0
        while c < c1 and row[width - c - 1].strip() == "":
            c += 1
        if c < c1:
            c1 = c
    return c1


###############################################################################
# FORMATTING
###############################################################################
def set_validation_list_by_cell(
        oCell: UnoCell, fields: List[str], default_string: Optional[str] = None,
        allow_blank: bool = False):
    oValidation = oCell.Validation
    oValidation.Type = uno.getConstantByName(
        "com.sun.star.sheet.ValidationType.LIST")
    oValidation.ShowErrorMessage = True
    oValidation.ShowList = uno.getConstantByName(
        "com.sun.star.sheet.TableValidationVisibility.UNSORTED")
    # TODO : improve me
    formula = ";".join('"' + a + '"' for a in fields)
    oValidation.setFormula1(formula)
    oValidation.IgnoreBlankCells = allow_blank
    oCell.Validation = oValidation

    if default_string is not None:
        oCell.String = default_string


def set_validation_list_by_name(
        oDoc: UnoSpreadsheet, cell_name: str, fields: List[str],
        default_string: Optional[str] = None, allow_blank: bool = False):
    oCell = get_named_cell(oDoc, cell_name)
    set_validation_list_by_cell(oCell, fields, default_string, allow_blank)


# CONDITIONAL

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


# from Andrew Pitonyak 5.14
# www.openoffice.org/documentation/HOW_TO/various_topics/AndrewMacro.odt
def find_or_create_number_format_style(oDoc: UnoSpreadsheet, fmt: str,
                                       locale: Optional[UnoStruct] = None
                                       ) -> int:
    oFormats = oDoc.getNumberFormats()
    if locale is None:
        oLocale = make_locale()
    else:
        oLocale = locale
    formatNum = oFormats.queryKey(fmt, oLocale, True)
    if formatNum == -1:
        formatNum = oFormats.addNew(fmt, oLocale)

    return formatNum


def create_filter(oRange: UnoRange):
    """
    Create a new filter
    @param oRange: the range to filter
    """
    oDoc = parent_doc(oRange)
    oDoc.CurrentController.select(oRange)
    oDispatchHelper = uno_service("com.sun.star.frame.DispatchHelper")
    oDispatchHelper.executeDispatch(oDoc.CurrentController.Frame,
                                    ".uno:DataFilterAutoFilter", "", 0, [])


def row_as_header(oHeaderRow: UnoRow):
    """
    Format the first row of the sheet
    @param oSheet:
    """
    oHeaderRow.CharWeight = FontWeight.BOLD
    oHeaderRow.CharWeightAsian = FontWeight.BOLD
    oHeaderRow.CharWeightComplex = FontWeight.BOLD
    oHeaderRow.IsTextWrapped = True
    oHeaderRow.OptimalHeight = True


def column_optimal_width(oColumn: UnoColumn, min_width: int = 2 * 1000,
                         max_width: int = 10 * 1000):
    """
    Sets the width of the column to an optimal value
    @param oColumn: the column
    @param min_width: the minimum width
    @param max_width: the maximum width
    """
    if oColumn.Width < min_width:
        oColumn.OptimalWidth = False
        oColumn.Width = min_width
    elif oColumn.Width > max_width:
        oColumn.OptimalWidth = False
        oColumn.IsTextWrapped = True
        oColumn.Width = max_width
    else:
        oColumn.OptimalWidth = True


def repeat_header_range(oSheet: UnoSheet, oRange: UnoRange):
    """
    Repeat the given range when printing.
    @param oSheet: the sheet
    @param oRange: the header range
    """
    used_range_address = get_used_range_address(oSheet)
    first_row_range_address = oRange.RangeAddress
    oSheet.setPrintAreas([used_range_address])
    oSheet.setPrintTitleRows(True)
    oSheet.setTitleRows(first_row_range_address)


###############################################################################
# MISC
###############################################################################
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


def copy_row_at_index(oSheet: UnoSheet, row: DATA_ROW, r: int):
    oRange = oSheet.getCellRangeByPosition(0, 0, len(row) - 1, 0)
    oRange.DataArray = row


def read_options_from_sheet_name(
        oDoc: UnoSpreadsheet, sheet_name: str,
        apply: Callable[[Tuple[str, str]], Tuple[str, str]] = lambda k_v: k_v):
    oSheet = oDoc.Sheets.getByName(sheet_name)
    oRangeAddress = get_used_range_address(oSheet)
    return read_options(oSheet, oRangeAddress, apply)


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
        :raises UnoRuntimeException: if Xray is not avaliable and
        `fail_on_error` is True.
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
        :@raises RuntimeException: if Xray is not avaliable and `fail_on_error`
         is True.
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
