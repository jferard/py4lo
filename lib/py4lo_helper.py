# Py4LO - Python Toolkit For LibreOffice Calc
#       Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
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
from pathlib import Path
from typing import (Any, Optional, List, cast, Callable, Mapping, Tuple,
                    Iterator, Union, Iterable)

from py4lo_commons import uno_path_to_url, CharProperties, Text, HTMLConverter
from py4lo_typing import (UnoSpreadsheetDocument, UnoController, UnoContext,
                          UnoService, UnoSheet, UnoRangeAddress, UnoRange,
                          UnoCell, UnoObject, DATA_ARRAY, UnoCellAddress,
                          UnoPropertyValue, DATA_ROW, UnoXScriptContext,
                          UnoColumn, UnoStruct, UnoEnum, UnoRow, DATA_VALUE,
                          UnoPropertyValues, UnoTextRange)

try:
    # noinspection PyUnresolvedReferences
    import unohelper

    # noinspection PyUnresolvedReferences
    import uno

    # noinspection PyUnresolvedReferences
    from com.sun.star.datatransfer import XTransferable


    class FrameSearchFlag:
        # noinspection PyUnresolvedReferences
        from com.sun.star.frame.FrameSearchFlag import (
            AUTO, PARENT, SELF, CHILDREN, CREATE, SIBLINGS, TASKS, ALL, GLOBAL)


    class BorderLineStyle:
        # noinspection PyUnresolvedReferences
        from com.sun.star.table.BorderLineStyle import (SOLID, )


    class ConditionOperator:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sheet.ConditionOperator import (FORMULA, )


    class FontWeight:
        # noinspection PyUnresolvedReferences
        from com.sun.star.awt.FontWeight import (BOLD, )


    class ValidationType:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sheet.ValidationType import (LIST, )


    class TableValidationVisibility:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sheet.TableValidationVisibility import (
            SORTEDASCENDING, UNSORTED)


    # noinspection PyUnresolvedReferences
    from com.sun.star.script.provider import ScriptFrameworkErrorException
    # noinspection PyUnresolvedReferences
    from com.sun.star.uno import (RuntimeException as UnoRuntimeException,
                                  Exception as UnoException)


    class PropertyState:
        # noinspection PyUnresolvedReferences
        from com.sun.star.beans.PropertyState import (
            AMBIGUOUS_VALUE, DIRECT_VALUE)

    class FontSlant:
        # noinspection PyUnresolvedReferences
        from com.sun.star.awt.FontSlant import (NONE, OBLIQUE, ITALIC)

except (ModuleNotFoundError, ImportError):
    from mock_constants import (  # noqa
        unohelper, uno, XTransferable, FrameSearchFlag, BorderLineStyle,
        ConditionOperator, FontWeight, ValidationType,
        TableValidationVisibility, ScriptFrameworkErrorException,
        UnoRuntimeException, UnoException, PropertyState, FontSlant
    )

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

    def __init__(self, doc: UnoSpreadsheetDocument, controller: UnoController,
                 frame, parent_win, script_provider,
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

        @return: the reflect service
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


def create_uno_service(sname: str, args: Optional[List[Any]] = None,
                       ctxt: Optional[UnoContext] = None) -> UnoService:
    sm = provider.service_manager
    if ctxt is None:
        return sm.createInstance(sname)
    else:
        if args is None:
            return sm.createInstanceWithContext(sname, ctxt)
        else:
            return sm.createInstanceWithArgumentsAndContext(sname, args, ctxt)


def create_uno_service_ctxt(sname: str,
                            args: Optional[List[Any]] = None) -> UnoService:
    return create_uno_service(sname, args, provider.ctxt)


# deprecated
uno_service = create_uno_service
# deprecated
uno_service_ctxt = create_uno_service_ctxt


def to_iter(o: UnoObject) -> Iterator[UnoObject]:
    """
    @param o: an XIndexAccess or XEnumerationAccession object
    @return: an iterator on `o`
    """
    try:
        count = o.Count
    except AttributeError:
        oEnum = o.createEnumeration()
        while oEnum.hasMoreElements():
            yield oEnum.nextElement()
    else:
        for i in range(count):
            yield o.getByIndex(i)


def to_enumerate(o: UnoObject) -> Iterator[Tuple[int, UnoObject]]:
    """
    @param o: an XIndexAccess or XEnumerationAccession object
    @return: an enumerate iterator on `o`
    """
    try:
        count = o.Count
    except AttributeError:
        oEnum = o.createEnumeration()
        i = 0
        while oEnum.hasMoreElements():
            yield i, oEnum.nextElement()
            i += 1
    else:
        for i in range(count):
            yield i, o.getByIndex(i)


def to_dict(oXNameAccess: UnoObject) -> Mapping[str, UnoObject]:
    return {
        name: oXNameAccess.getByName(name)
        for name in oXNameAccess.getElementNames()
    }


def to_items(oXNameAccess: UnoObject) -> Iterator[Tuple[str, UnoObject]]:
    return (
        (name, oXNameAccess.getByName(name))
        for name in oXNameAccess.ElementNames
    )


def remove_all(oXAccess: UnoObject):
    try:
        for name in oXAccess.ElementNames:
            oXAccess.removeByName(name)
    except AttributeError:
        while oXAccess.Count:
            oXAccess.removeByIndex(0)


def parent_doc(oRange: UnoRange) -> UnoSpreadsheetDocument:
    """
    Find the document that owns this range.

    @param oRange: the range (range, sheet, cell)
    @return: the document to which this range belongs
    """
    oSheet = oRange.Spreadsheet
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


def get_named_cells(oDoc: UnoSpreadsheetDocument, name: str) -> UnoRange:
    return oDoc.NamedRanges.getByName(name).ReferredCells


def get_named_cell(oDoc: UnoSpreadsheetDocument, name: str) -> UnoCell:
    return get_named_cells(oDoc, name).getCellByPosition(0, 0)


def get_main_cell(oCell: UnoCell) -> UnoCell:
    """
    Return the main cell.
    :param oCell: a cell inside a merged cells group
    :return: the main cell
    """
    oSheet = oCell.Spreadsheet
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.collapseToMergedArea()
    return oSheet.getCellByPosition(oCursor.RangeAddress.StartColumn,
                                    oCursor.RangeAddress.StartRow)


##############################################################################
# STRUCTS
##############################################################################

def create_uno_struct(struct_id: str, **kwargs):
    struct = uno.createUnoStruct(struct_id)
    for k, v in kwargs.items():
        struct.__setattr__(k, v)
    return struct


make_struct = create_uno_struct


def make_pv(name: str, value: Any) -> UnoPropertyValue:
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


def make_pvs(d: Optional[Mapping[str, Any]] = None
             ) -> Tuple[UnoPropertyValue, ...]:
    if d is None:
        return tuple()
    else:
        return tuple(make_pv(n, v) for n, v in d.items())


def update_pvs(pvs: Iterable[UnoPropertyValue], d: Mapping[str, Any]):
    """
    Update in place some of the pvs, based on names
    @param pvs: the pvs
    @param d: the mapping name to value
    """
    for pv in pvs:
        if pv.Name in d:
            pv.Value = d[pv.Name]


def make_locale(language: str = "", region: str = "",
                subtags: Optional[List[str]] = None) -> UnoStruct:
    """
    Create a locale

    @param region: ISO 3166 Country Code.
    @param language: ISO 639 Language Code.
    @param subtags: BCP 47
    @return: the locale
    """
    locale = uno.createUnoStruct('com.sun.star.lang.Locale')
    if not subtags:
        locale.Country = region
        locale.Language = language
    else:
        locale.Language = "qlt"
        if region:
            locale.Variant = "-".join([language, region] + subtags)
        else:
            locale.Variant = "-".join([language] + subtags)
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


def make_sort_field(field_position: int, asc: bool = True):
    sf = uno.createUnoStruct('com.sun.star.table.TableSortField')
    sf.Field = field_position
    sf.IsAscending = asc
    return sf


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


def copy_range(oSourceRange: UnoRange):
    """
    Copy a range to the clipboard
    :param oSourceRange the range (may be a sheet)
    """
    oSourceDoc = parent_doc(oSourceRange)
    oSourceController = oSourceDoc.CurrentController
    oSourceController.select(oSourceRange)
    provider.dispatcher.executeDispatch(
        oSourceController, ".uno:Copy", "", 0, [])
    # unselect
    oRanges = oSourceDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
    oSourceController.select(oRanges)


def paste_range(oDestSheet: UnoSheet, oDestAddress: UnoCellAddress,
                formulas: bool = False):
    """
    """
    oDestDoc = parent_doc(oDestSheet)
    oDestController = oDestDoc.CurrentController
    oDestCell = oDestSheet.getCellByPosition(oDestAddress.Column,
                                             oDestAddress.Row)
    oDestController.select(oDestCell)
    if formulas:
        provider.dispatcher.executeDispatch(
            oDestController, ".uno:Paste", "", 0, [])
    else:
        # TODO: propose more options
        args = make_pvs({
            "Flags": "SVDT", "FormulaCommand": 0, "SkipEmptyCells": False,
            "Transpose": False, "AsLink": False, "MoveMode": 4
        })
        provider.dispatcher.executeDispatch(
            oDestController, ".uno:InsertContents", "", 0, args)
    # unselect
    oRanges = oDestDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
    oDestController.select(oRanges)


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
        count = top_void_row_count(data_array)
        if count == len(data_array):
            return None

        start_row += count
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
    @return: the number of void row at the bottom
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
    @return: the number of void row at the left
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
    @return: the number of void row at the right
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
        oCell: UnoCell, values: List[Any],
        default_string: Optional[str] = None,
        ignore_blank: bool = True, sorted_values: bool = False,
        show_error: bool = True):
    factory = ValidationFactory().list().values(values)
    factory.ignore_blank(ignore_blank)
    factory.sort_values(sorted_values)
    factory.show_error(show_error)
    factory.update(oCell)

    if default_string is not None:
        oCell.String = default_string


class ValidationFactory:
    def list(self):
        return ListValidationBuilder()


class ListValidationBuilder:
    def __init__(self):
        self._values = []
        self._default_string = None
        self._ignore_blank = False
        self._sort_values = False
        self._show_error = True

    def values(self, values: List[Any]) -> "ListValidationBuilder":
        self._values = values
        return self

    def default_string(self, default_string: str) -> "ListValidationBuilder":
        self._default_string = default_string
        return self

    def ignore_blank(self, ignore_blank: bool) -> "ListValidationBuilder":
        self._ignore_blank = ignore_blank
        return self

    def sort_values(self, sorted_values: bool) -> "ListValidationBuilder":
        self._sort_values = sorted_values
        return self

    def show_error(self, show_error: bool) -> "ListValidationBuilder":
        self._show_error = show_error
        return self

    def update(self, oCell: UnoCell):
        oValidation = oCell.Validation
        oValidation.Type = ValidationType.LIST
        oValidation.IgnoreBlankCells = self._ignore_blank
        if self._sort_values:
            oValidation.ShowList = TableValidationVisibility.UNSORTED
        else:
            oValidation.ShowList = TableValidationVisibility.SORTEDASCENDING
        oValidation.ShowErrorMessage = self._show_error

        oValidation.Formula1 = ";".join(
            quote_element(f) for f in self._values)
        oCell.Validation = oValidation


def quote_element(value: Any) -> str:
    """
    Quote a list element (see formula).
    TODO: use a locale

    :param value: the value
    :return: the quoted value
    """
    if isinstance(value, str):
        value = value.replace('"', '\\"')
    elif isinstance(value, float):
        pass  # if locale: replace(".", ",")
    # elif isinstance(value, bool):
    #     pass  # if locale
    # elif isinstance(value, (date, datetime)):
    #     pass # if locale
    return '"{}"'.format(value)


def sort_range(oRange: UnoRange, sort_fields: Tuple[UnoStruct, ...],
               has_header: bool = True):
    """
    @param oRange:
    @param sort_fields:
    @param has_header: True if the range has a header
    @return:
    """
    typed_sort_fields = uno.Any('[]com.sun.star.table.TableSortField',
                                sort_fields)
    sort_descriptor = oRange.createSortDescriptor()
    update_pvs(sort_descriptor,
               {'ContainsHeader': has_header, 'SortFields': typed_sort_fields})
    oRange.sort(sort_descriptor)


# CONDITIONAL
def clear_conditional_format(oColoredColumns: UnoRange):
    oConditionalFormat = oColoredColumns.ConditionalFormat
    oConditionalFormat.clear()


def conditional_format_on_formulas(
        oColoredColumns: UnoRange, style_by_formula: Mapping[str, str],
        oSrcAddress: UnoCellAddress):
    oConditionalFormat = oColoredColumns.ConditionalFormat
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
def find_or_create_number_format_style(oDoc: UnoSpreadsheetDocument, fmt: str,
                                       locale: Optional[UnoStruct] = None
                                       ) -> int:
    oFormats = oDoc.NumberFormats
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
    oController = oDoc.CurrentController
    oController.select(oRange)
    provider.dispatcher.executeDispatch(
        oController.Frame, ".uno:DataFilterAutoFilter", "", 0, [])
    # unselect
    oRanges = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
    oController.select(oRanges)


def remove_filter(oRange: UnoRange):
    """
    Remove the existing filter on a range
    @param oRange: The range
    """
    # True means "empty"
    oFilterDescriptor = oRange.createFilterDescriptor(True)
    oRange.filter(oFilterDescriptor)
    create_filter(oRange)


def row_as_header(oHeaderRow: UnoRow):
    """
    Format the first row of the sheet
    @param oHeaderRow:
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


def set_print_area(oSheet: UnoSheet, oTitleRows: Optional[UnoRange] = None):
    """
    Repeat the given range when printing.
    @param oSheet: the sheet
    @param oTitleRows: the header range
    """
    used_range_address = get_used_range_address(oSheet)
    oSheet.setPrintAreas([used_range_address])
    if oTitleRows is not None:
        title_rows_address = oTitleRows.RangeAddress
        oSheet.setPrintTitleRows(True)
        oSheet.setTitleRows(title_rows_address)


A3_LARGE = 42000
A4_LARGE = A3_SMALL = 29700
A4_SMALL = 21000


def get_page_style(oSheet: UnoSheet) -> UnoService:
    """
    @param oSheet: a sheet
    @return: the page style of this sheet
    """
    page_style_name = oSheet.PageStyle
    oDoc = parent_doc(oSheet)
    oStyle = oDoc.StyleFamilies.getByName("PageStyles").getByName(
        page_style_name)
    return oStyle


def set_paper(oSheet: UnoSheet):
    """
    Set the paper for this sheet
    @param oSheet: the sheet
    """
    oPageStyle = get_page_style(oSheet)
    size = get_used_range(oSheet).Size
    set_paper_to_size(oPageStyle, size)


def set_paper_to_size(oPageStyle: UnoService, size: UnoStruct):
    """
    Make the paper of this style match this area.
    @param oPageStyle: the page style
    @param size: the size of the area
    """
    if size.Height >= size.Width:  # prefer portrait
        oPageStyle.IsLandscape = False
        oPageStyle.ScaleToPagesX = 1
        oPageStyle.ScaleToPagesY = 0
        if size.Width > A4_SMALL:  # too wide for A4
            style_size = create_uno_struct("com.sun.star.awt.Size",
                                           Width=A3_SMALL, Height=A3_LARGE)
        else:
            style_size = create_uno_struct("com.sun.star.awt.Size",
                                           Width=A4_SMALL, Height=A4_LARGE)
    else:  # prefer landscape
        oPageStyle.IsLandscape = True
        oPageStyle.ScaleToPagesX = 0
        oPageStyle.ScaleToPagesY = 1
        if size.Height > A4_SMALL:  # too high for A4
            style_size = create_uno_struct("com.sun.star.awt.Size",
                                           Width=A3_LARGE, Height=A3_SMALL)
        else:
            style_size = create_uno_struct("com.sun.star.awt.Size",
                                           Width=A4_LARGE, Height=A4_SMALL)
    oPageStyle.Size = style_size


def add_link(oCell: UnoCell, text: str, url: str, wrap_at: int = -1):
    """
    Add a link to the end of this cell.
    @param oCell: the cell
    @param text: the text
    @param url: the url of the link
    @param wrap_at: if -1, don't wrap, else wrap the text every n chars
    """
    oCursor = oCell.Text.createTextCursorByRange(oCell.Text.End)
    oDoc = parent_doc(oCell)

    if wrap_at == -1:
        lines = [text]
    else:
        lines = _wrap_text(text, wrap_at)

    for line in lines:
        text_field = oDoc.createInstance("com.sun.star.text.TextField.URL")
        text_field.Representation = line
        text_field.URL = url
        oCell.insertTextContent(oCursor, text_field, False)


def _wrap_text(text: str, wrap_at: int):
    lines = []
    words = text.split()
    word = words[0]
    cur = [word]
    c = len(word) + 1
    for word in words[1:]:
        cur_len = len(word)
        if c + cur_len > wrap_at:
            lines.append(" ".join(cur))
            cur = [word]
            c = cur_len + 1
        else:
            cur.append(word)
            c += cur_len
    lines.append(" ".join(cur))
    return lines


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


def open_in_calc(filename: Union[str, Path], target: str = Target.BLANK,
                 frame_flags=FrameSearchFlag.AUTO,
                 **kwargs) -> UnoSpreadsheetDocument:
    """
    Open a document in calc
    :param filename: the name of the file to open
    :param target: "the name of the frame to view the document in" or a special
    target
    :param frame_flags: where to search the frame
    :param kwargs: les paramètres d'ouverture
    :return: a reference on the doc
    """
    url = uno_path_to_url(filename)
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
        pvs: Optional[UnoPropertyValues] = None
) -> "DocBuilder":
    if pvs is None:
        pvs = tuple()
    return DocBuilder(url, taget_frame_name, search_flags, pvs)


def new_doc(url: NewDocumentUrl = NewDocumentUrl.Calc,
            taget_frame_name: Target = Target.BLANK,
            search_flags: FrameSearchFlag = FrameSearchFlag.AUTO,
            pvs: Optional[UnoPropertyValues] = None) -> UnoSpreadsheetDocument:
    """Create a blank new doc"""
    return doc_builder(url, taget_frame_name, search_flags, pvs).build()


class DocBuilder:
    """
    Todo: store and then build
    """

    def __init__(self, url: NewDocumentUrl, taget_frame_name: Target,
                 search_flags: FrameSearchFlag,
                 pvs: UnoPropertyValues):
        """Create a blank new doc"""
        self._oDoc = provider.desktop.loadComponentFromURL(
            url, taget_frame_name, search_flags, pvs)
        self._oDoc.lockControllers()

    def build(self) -> UnoSpreadsheetDocument:
        self._oDoc.unlockControllers()
        return self._oDoc

    def sheet_names(self, sheet_names: List[str],
                    expand_if_necessary: bool = True,
                    trunc_if_necessary: bool = True) -> "DocBuilder":
        oSheets = self._oDoc.Sheets
        it = iter(sheet_names)
        s = 0

        initial_count = oSheets.Count
        try:
            # rename
            while s < initial_count:
                oSheet = oSheets.getByIndex(s)
                oSheet.Name = next(it)  # may raise a StopIteration
                s += 1

            if s != initial_count:
                raise AssertionError("s={} vs oSheets.getCount()={}".format(
                    s, initial_count))

            if expand_if_necessary:
                # add
                for sheet_name in it:
                    oSheets.insertNewByName(sheet_name, s)
                    s += 1
        except StopIteration:  # it
            if s > initial_count:
                raise AssertionError("s={} vs oSheets.getCount()={}".format(
                    s, oSheets.getCount()))
            if trunc_if_necessary:
                self.trunc_to_count(s)

        return self

    def apply_func_to_sheets(
            self, func: Callable[[UnoSheet], None]) -> "DocBuilder":
        oSheets = self._oDoc.Sheets
        for oSheet in to_iter(oSheets):
            func(oSheet)
        return self

    def apply_func_list_to_sheets(
            self, funcs: List[Callable[[UnoSheet], None]]) -> "DocBuilder":
        oSheets = self._oDoc.Sheets
        for func, oSheet in zip(funcs, to_iter(oSheets)):
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

    def make_base_sheet(self, func: Callable[[UnoSheet], None]
                        ) -> "DocBuilder":
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
        while final_sheet_count < oSheets.Count:
            oSheet = oSheets.getByIndex(final_sheet_count)
            oSheets.removeByName(oSheet.Name)

        return self


###############################################################################
# MISC
###############################################################################
def read_options(oSheet: UnoSpreadsheetDocument, aAddress: UnoRangeAddress,
                 apply: Union[
                     Callable[[str, str], Tuple[str, str]],
                     Callable[[str, List[str]], Tuple[str, List[str]]]
                 ] = lambda k_v: k_v
                 ) -> Mapping[str, DATA_VALUE]:
    options = {}
    if aAddress.StartColumn == aAddress.EndColumn:
        return {}

    oRange = oSheet.getCellRangeByPosition(
        aAddress.StartColumn, aAddress.StartRow,
        aAddress.EndColumn, aAddress.EndRow)
    data_array = oRange.DataArray
    for row in data_array:
        k = row[0]
        v = rtrim_row(row[1:])
        k, v = apply(k, v)
        if not k:
            continue
        options[k] = v
    return options


def rtrim_row(row: DATA_ROW, null="") -> Union[DATA_ROW, DATA_VALUE]:
    if len(row) == 0:
        return null

    i = len(row)
    while i > 0:
        v = row[i - 1]
        if v != "":
            break
        i -= 1

    if i == 0:
        return null
    elif i == 1:
        return row[0]
    else:
        return row[:i]


def read_options_from_sheet_name(
        oDoc: UnoSpreadsheetDocument, sheet_name: str,
        apply: Union[
            Callable[[str, str], Tuple[str, str]],
            Callable[[str, List[str]], Tuple[str, List[str]]]
        ] = lambda k_v: k_v):
    oSheet = oDoc.Sheets.getByName(sheet_name)
    oRangeAddress = get_used_range_address(oSheet)
    return read_options(oSheet, oRangeAddress, apply)


def copy_row_at_index(oSheet: UnoSheet, row: DATA_ROW, r: int):
    oRange = oSheet.getCellRangeByPosition(0, r, len(row) - 1, r)
    oRange.DataArray = row


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
            self._ignore_xray = True
            if fail_on_error:
                raise UnoRuntimeException(
                    "\nBasic library Xray is not installed",
                    self._provider.ctxt)

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
                self._oMRI = create_uno_service("mytools.Mri")
            except UnoException:
                self._ignore_mri = True
                if fail_on_error:
                    raise UnoRuntimeException("\nMRI is not installed",
                                              self._provider.ctxt)
        if self._ignore_mri:
            return

        self._oMRI.inspect(obj)


TEXT_FLAVOR = ("text/plain;charset=utf-16", "Unicode-text")
HTML_FLAVOR = ("text/html;charset=utf-8", "HTML")


def copy_to_clipboard(value: Any, flavor: Tuple[str, str] = TEXT_FLAVOR):
    """See https://forum.openoffice.org/en/forum/viewtopic.php?t=93562"""
    oClipboard = create_uno_service(
        "com.sun.star.datatransfer.clipboard.SystemClipboard")
    oClipboard.setContents(Transferable(value, flavor), None)


def get_from_clipboard(flavor: Tuple[str, str] = TEXT_FLAVOR) -> Optional[Any]:
    """See https://forum.openoffice.org/en/forum/viewtopic.php?t=93562"""
    oClipboard = create_uno_service(
        "com.sun.star.datatransfer.clipboard.SystemClipboard")
    oContents = oClipboard.getContents()
    oTypes = oContents.getTransferDataFlavors()

    for oType in oTypes:
        if oType.MimeType == flavor[0]:
            return oContents.getTransferData(oType)

    return None


class Transferable(unohelper.Base, XTransferable):
    def __init__(self, value: Any, flavor: Tuple[str, str]):
        self._value = value
        self._flavor = flavor

    def getTransferData(self, aFlavor: UnoObject):
        if aFlavor.MimeType == self._flavor[0]:
            return self._value

    def getTransferDataFlavors(self):
        flavor = create_uno_struct("com.sun.star.datatransfer.DataFlavor",
                                   MimeType=self._flavor[0],
                                   HumanPresentableName=self._flavor[1])
        return [flavor]

    def isDataFlavorSupported(self, aFlavor: UnoObject) -> bool:
        return aFlavor.MimeType == self._flavor[0]


def char_iter(oXSimpleText) -> Iterator[UnoTextRange]:
    """
    Iterator over the chars of a text.
    Beware: the cursor is always the same.

    @param oXSimpleText: the text
    @return: the iterator
    """
    oCursor = oXSimpleText.createTextCursor()
    oCursor.gotoStart(False)
    while oCursor.goRight(1, True):
        yield oCursor
        oCursor.goRight(0, False)


def char_properties_from_uno_text_range(
        text_range: UnoTextRange) -> CharProperties:
    """Create a new CharProperties object from a text range"""
    italic = text_range.CharPosture == FontSlant.OBLIQUE or text_range.CharPosture == FontSlant.ITALIC
    script = None
    if text_range.CharEscapementHeight < 100:
        if text_range.CharEscapement < 0:
            script = "sub"
        elif text_range.CharEscapement > 0:
            script = "sup"
    return CharProperties(
        text_range.CharFontName, text_range.CharHeight, text_range.CharWeight,
        italic, text_range.CharBackColor, text_range.CharColor,
        text_range.CharOverline, text_range.CharStrikeout,
        text_range.CharUnderline, script
    )


def text_from_uno_text_range(text_range: UnoTextRange) -> Text:
    """Create a new Text object from a text range"""
    return Text(
        text_range.String,
        char_properties_from_uno_text_range(text_range)
    )


def convert_to_html(text_range: UnoTextRange) -> str:
    uno_chars = char_iter(text_range)
    return HTMLConverter().convert(
        [text_from_uno_text_range(uno_char) for uno_char in uno_chars])
