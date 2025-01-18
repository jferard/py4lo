# Py4LO - Python Toolkit For LibreOffice Calc
#       Copyright (C) 2016-2024 J. Férard <https://github.com/jferard>
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
"""
The module py4lo_helper deals with LO objects.

It provides a lot of simple functions to handle those objects.
"""
import datetime as dt
# mypy: disable-error-code="import-untyped,import-not-found"
import logging
from contextlib import contextmanager
from enum import Enum
from locale import getlocale
from pathlib import Path
from typing import (Any, Optional, List, cast, Callable, Mapping, Tuple,
                    Iterator, Union, Iterable, ContextManager)

from py4lo_commons import uno_path_to_url
from py4lo_typing import (UnoSpreadsheetDocument, UnoController, UnoContext,
                          UnoService, UnoSheet, UnoRangeAddress, UnoRange,
                          UnoCell, UnoObject, DATA_ARRAY, UnoCellAddress,
                          UnoPropertyValue, DATA_ROW, UnoXScriptContext,
                          UnoColumn, UnoStruct, UnoEnum, UnoRow, DATA_VALUE,
                          UnoPropertyValues, UnoTextRange, lazy, UnoControl,
                          UnoDispatcher, UnoDesktop)

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

except ImportError:
    from _mock_constants import (  # type:ignore[assignment]
        BorderLineStyle,  # pyright: ignore[reportGeneralTypeIssues]
        ConditionOperator,  # pyright: ignore[reportGeneralTypeIssues]
        FontSlant,  # pyright: ignore[reportGeneralTypeIssues]
        FontWeight,  # pyright: ignore[reportGeneralTypeIssues]
        FrameSearchFlag,  # pyright: ignore[reportGeneralTypeIssues]
        PropertyState,  # pyright: ignore[reportGeneralTypeIssues]
        TableValidationVisibility,  # pyright: ignore[reportGeneralTypeIssues]
        ValidationType,  # pyright: ignore[reportGeneralTypeIssues]
    )
    from _mock_objects import (  # type: ignore[assignment]
        ScriptFrameworkErrorException, \
        # pyright: ignore[reportGeneralTypeIssues]
        UnoException,  # pyright: ignore[reportGeneralTypeIssues]
        UnoRuntimeException,  # pyright: ignore[reportGeneralTypeIssues]
        XTransferable,  # pyright: ignore[reportGeneralTypeIssues]
        uno,  # pyright: ignore[reportGeneralTypeIssues]
        unohelper,  # pyright: ignore[reportGeneralTypeIssues]
    )

###############################################################################
# BASE
###############################################################################
provider = cast(Optional["_ObjectProvider"], None)
_inspector = cast(Optional["_Inspector"], None)
xray = cast(Optional[Callable], None)
mri = cast(Optional[Callable], None)


def init(xsc: UnoXScriptContext):
    """
    Mandatory call from entry script with XSCRIPTCONTEXT as argument.
    Initialize the provider, the `xray` and `mri` functions.

    ```
    import py4lo_helper
    from py4lo helper import provider

    py4lo_helper.init(XSCRIPTCONTEXT)
    xray(provider.doc) # let's hope xray is installed...
    ...
    ```

    @param xsc: XSCRIPTCONTEXT
    """
    global provider, _inspector, xray, mri
    provider = _ObjectProvider.create(xsc)
    _inspector = _Inspector(provider)
    xray = _inspector.xray
    mri = _inspector.mri


def init_from_component_context(component_ctxt: UnoContext):
    """
    To use Py4LO inside an extension (see XJobExecutor)

    @param component_ctxt: the component context parameter for the JobExecutor
    """
    global provider, _inspector, xray, mri
    provider = _ObjectProvider.create_from_component_context(component_ctxt)
    _inspector = _Inspector(provider)
    xray = _inspector.xray
    mri = _inspector.mri


class _ObjectProvider:
    """
    A lazy object provider. Available objects:

    * provider.doc: the document (ThisComponent in Basic,
    see: com.sun.star.frame.XModel)
    * provider.controller: the current controller of the document (see: com.sun.star.frame.XController)
    * provider.frame: the frame of the controller (see: com.sun.star.frame.XFrame)
    * provider.parent_win: the container window of the frame (see: com.sun.star.awt.XWindow)
    * provider.script_provider: the script provider (see: com.sun.star.script.provider.XScriptProvider)
    * provider.ctxt: the component context (see: com.sun.star.uno.XComponentContext)
    * provider.service_manager: the service manager (see: com.sun.star.lang.XMultiComponentFactory)
    * provider.desktop: the desktop (see: com.sun.star.frame.XDesktop)

    Lazy:
    * provider.reflect: the reflect service (see: com.sun.star.reflection.CoreReflection)
    * provider.dispatcher: the dispatcher (see: com.sun.star.frame.DispatchHelper)
    """

    @staticmethod
    def create(xsc: UnoXScriptContext) -> "_ObjectProvider":
        """
        Create a new _ObjectProvider

        @param xsc: XSCRIPTCONTEXT
        @return: the object provider
        """
        doc = cast(UnoSpreadsheetDocument, xsc.getDocument())
        controller = doc.CurrentController
        frame = controller.Frame
        parent_win = frame.ContainerWindow
        script_provider = doc.getScriptProvider()
        ctxt = xsc.getComponentContext()
        service_manager = ctxt.getServiceManager()
        desktop = xsc.getDesktop()
        return _ObjectProvider(doc, controller, frame, parent_win,
                               script_provider, ctxt, service_manager, desktop)

    @staticmethod
    def create_from_component_context(
            component_ctxt: UnoContext) -> "_ObjectProvider":
        """
        Create a new _ObjectProvider

        @param component_ctxt: the component context
        @return: the object provider
        """
        desktop = component_ctxt.getByName(
            "/singletons/com.sun.star.frame.theDesktop")
        doc = cast(UnoSpreadsheetDocument, desktop.getCurrentComponent())
        controller = doc.CurrentController
        frame = controller.Frame
        parent_win = frame.ContainerWindow
        script_provider = doc.getScriptProvider()
        service_manager = component_ctxt.getServiceManager()
        return _ObjectProvider(
            doc, controller, frame, parent_win,
            script_provider, component_ctxt, service_manager, desktop)

    def __init__(self, doc: UnoSpreadsheetDocument, controller: UnoController,
                 frame, parent_win, script_provider,
                 ctxt: UnoContext, service_manager, desktop: UnoDesktop):
        """
        @param doc: the document (ThisComponent in Basic, see: com.sun.star.frame.XModel)
        @param controller: the current controller of the document (see: com.sun.star.frame.XController)
        @param frame: the frame of the controller (see: com.sun.star.frame.XFrame)
        @param parent_win: the container window of the frame (see: com.sun.star.awt.XWindow)
        @param script_provider: the script provider (see: com.sun.star.script.provider.XScriptProvider)
        @param ctxt: the component context (see: com.sun.star.uno.XComponentContext)
        @param service_manager: the service manager (see: com.sun.star.lang.XMultiComponentFactory)
        @param desktop: the desktop (see: com.sun.star.frame.XDesktop)
        """
        self.doc = doc
        self.controller = controller
        self.frame = frame
        self.parent_win = parent_win
        self.script_provider = script_provider
        self.ctxt = ctxt
        self.service_manager = service_manager
        self.desktop = desktop
        self._script_provider_factory = lazy(UnoService)
        self._script_provider = lazy(UnoService)
        self._reflect = lazy(UnoService)
        self._dispatcher = lazy(UnoDispatcher)

    def get_script_provider_factory(self) -> UnoService:
        """
        This service is used to create MasterScriptProviders.

        @return: a com.sun.star.script.provider.MasterScriptProviderFactory
        """
        if self._script_provider_factory is None:
            self._script_provider_factory = cast(
                UnoService,
                self.service_manager.createInstanceWithContext(
                    "com.sun.star.script.provider.MasterScriptProviderFactory",
                    self.ctxt)
            )
        return self._script_provider_factory

    def get_script_provider(self) -> UnoService:
        """
        This interface provides a factory for obtaining objects implementing
        the XScript interface

        @return: a com.sun.star.script.provider.XScriptProvider
        """
        if self._script_provider is None:
            self._script_provider = cast(
                UnoService,
                self.get_script_provider_factory().createScriptProvider("")
            )
        return self._script_provider

    @property
    def reflect(self) -> UnoService:
        """
        This service is the implementation of the reflection API

        @return: the reflect service (see: com.sun.star.reflection.CoreReflection)
        """
        if self._reflect is None:
            self._reflect = cast(
                UnoService,
                self.service_manager.createInstance(
                    "com.sun.star.reflection.CoreReflection")
            )
        return self._reflect

    @property
    def dispatcher(self) -> UnoDispatcher:
        """
        provides an easy way to dispatch a URL using one call instead of
        multiple ones.

        @return: the dispatcher (see: com.sun.star.frame.DispatchHelper)
        """
        if self._dispatcher is None:
            self._dispatcher = cast(
                UnoDispatcher,
                self.service_manager.createInstance(
                    "com.sun.star.frame.DispatchHelper")
            )
        return self._dispatcher

    def get_dialog(self, dialog_name: str,
                   dialog_location: str = "document") -> UnoControl:
        """
        Get a GUI existing dialog

        @param dialog_name: the name of the dialog
        @param dialog_location: the location of the dialog
        @return: the dialog (see: com.sun.star.awt.UnoControlDialog)
        """
        oDialogProvider = self.service_manager.createInstanceWithArguments(
            "com.sun.star.awt.DialogProvider", (self.doc,))
        dialog_url = "vnd.sun.star.script:{}?location={}".format(
            dialog_name, dialog_location)
        oDialog = oDialogProvider.createDialog(dialog_url)
        return oDialog


def get_provider() -> _ObjectProvider:
    """For use in other modules, inside an LibreOffice extension"""
    return provider


def create_uno_service(sname: str, args: Optional[List[Any]] = None,
                       ctxt: Optional[UnoContext] = None) -> UnoService:
    """
    Create a new UNO service
    @param sname: the service name
    @param args: optional args
    @param ctxt: optional context
    @return: the service
    """
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
    """
    Create a new UNO service with the current context
    @param sname: the service name
    @param args: optional args
    @return: the service
    """
    return create_uno_service(sname, args, provider.ctxt)


# deprecated
uno_service = create_uno_service
# deprecated
uno_service_ctxt = create_uno_service_ctxt


def to_iter(o: UnoObject) -> Iterator[UnoObject]:
    """
    Return an iterator from a com.sun.star.container.XIndexAccess or a
    com.sun.star.container.XEnumerationAccess.

    Example:
    ```
    for oSheet in to_iter(oSheets):
        ...
    ```

    @param o: an XIndexAccess or XEnumerationAccession object
    @return: an iterator
    """
    try:
        count = o.Count
        for i in range(count):
            yield o.getByIndex(i)
    except AttributeError:
        try:
            oEnum = o.createEnumeration()
            while oEnum.hasMoreElements():
                yield oEnum.nextElement()
        except AttributeError:
            raise TypeError(repr(o))


def to_enumerate(o: UnoObject) -> Iterator[Tuple[int, UnoObject]]:
    """
    Return an iterator from a com.sun.star.container.XIndexAccess or a
    com.sun.star.container.XEnumerationAccess object.

    Example:
    ```
    for i, oSheet in to_enumerate(oSheets):
        ...
    ```

    @param o: an XIndexAccess or XEnumerationAccession object
    @return: an enumerate iterator (number, value)
    """
    try:
        count = o.Count  # type: ignore[union-attr]
        for i in range(count):
            yield i, o.getByIndex(i)  # type: ignore[union-attr]
    except AttributeError:
        try:
            oEnum = o.createEnumeration()  # type: ignore[union-attr]
            i = 0
            while oEnum.hasMoreElements():
                yield i, oEnum.nextElement()
                i += 1
        except AttributeError:
            raise TypeError(repr(o))


def to_dict(oXNameAccess: UnoObject) -> Mapping[str, UnoObject]:
    """
    Return a dictionary from a com.sun.star.container.XNameAccess object.

    Example:
    ```
    oSheet_by_name = to_dict(oSheets)
    ```

    @param oXNameAccess: an XNameAccess object
    @return: a mapping
    """
    try:
        return {
            name: oXNameAccess.getByName(name)
            for name in oXNameAccess.ElementNames
        }
    except AttributeError as e:
        raise TypeError(repr(oXNameAccess) + str(e))


def to_items(oXNameAccess: UnoObject) -> Iterator[Tuple[str, UnoObject]]:
    """
    Return a dictionary from a com.sun.star.container.XNameAccess object.

    Example:
    ```
    for name, oSheet in to_enumerate(oSheets):
        ...
    ```

    @param oXNameAccess: an XNameAccess object
    @return: a mapping
    """
    try:
        return (
            (name, oXNameAccess.getByName(name))
            for name in oXNameAccess.ElementNames
        )
    except AttributeError:
        raise TypeError(repr(oXNameAccess))


def remove_all(o: UnoObject):
    """
    Remove all elements from a container (com.sun.star.container.XIndexContainer
    or com.sun.star.container.XNameContainer)

    @param o: the XIndexContainer or XNameContainer object
    """
    try:
        while o.Count:  # type: ignore[union-attr]
            o.removeByIndex(0)  # type: ignore[union-attr]
    except AttributeError:
        try:
            for name in o.ElementNames:  # type: ignore[union-attr]
                o.removeByName(name)  # type: ignore[union-attr]
        except AttributeError:
            raise TypeError(repr(o))


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
    """
    @param oDoc: the document
    @param name: the name of the range
    @return: the range itself (see: com.sun.star.table.XCellRange)
    """
    return oDoc.NamedRanges.getByName(name).ReferredCells


def get_named_cell(oDoc: UnoSpreadsheetDocument, name: str) -> UnoCell:
    """
    @param oDoc: the document
    @param name: the name of the cell
    @return: the first cell of the range (see: com.sun.star.table.XCell)
    """
    return get_named_cells(oDoc, name).getCellByPosition(0, 0)


def get_main_cell(oCell: UnoCell) -> UnoCell:
    """
    Return the main cell of the group of merged cells containing
    this cell.

    @param oCell: a cell inside a merged cells group (see: com.sun.star.table.XCell)
    @return: the main cell (see: com.sun.star.table.XCell)
    """
    oSheet = oCell.Spreadsheet
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.collapseToMergedArea()
    return oSheet.getCellByPosition(oCursor.RangeAddress.StartColumn,
                                    oCursor.RangeAddress.StartRow)


##############################################################################
# STRUCTS
##############################################################################

def create_uno_struct(struct_id: str, **kwargs) -> UnoStruct:
    """
    Create a new UNO struct

    Example:
    ```
    oDate = create_uno_struct("com.sun.star.util.Date", Day=26, Month=11, Year=2024)
    ```

    For PropertyValue, see make_pv or make_full_pv.

    @param struct_id: the id of the sructure
    @param kwargs: the parameters
    @return: the UNO struct
    """
    struct = uno.createUnoStruct(struct_id)
    for k, v in kwargs.items():
        struct.__setattr__(k, v)
    return struct


make_struct = create_uno_struct


def make_pv(name: str, value: Any) -> UnoPropertyValue:
    """
    Create a com.sun.star.beans.PropertyValue object

    @param name: the name of the PropertyValue
    @param value: the value of the PropertyValue
    @return: the PropertyValue
    """
    pv = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    pv.Name = name
    pv.Value = value
    pv.Handle = 0
    pv.State = PropertyState.DIRECT_VALUE
    return pv


def make_full_pv(name: str, value: str, handle: int = -1,
                 state: Optional[UnoEnum] = None) -> UnoPropertyValue:
    """
    Create a com.sun.star.beans.PropertyValue object with handle and state

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
    """
    Create a tuple of PropertyValue objects from a dictionary.
    Each key, value pair will be mapped to a PropertyValue having the key as
    `Name` and the value as `Value`.

    Example:
    ```
    pvs = make_pvs({"FilterName": Filter.CSV, "FilterOptions": "...",
                    "Hidden": True})
    ```

    @param d: the dictionary
    @return: the tuple of PropertyValue objects
    """
    if d is None:
        return tuple()
    else:
        return tuple(make_pv(n, v) for n, v in d.items())


def update_pvs(pvs: Iterable[UnoPropertyValue], d: Mapping[str, Any]):
    """
    Update in place some of the pvs, based on names.

    @param pvs: the pvs
    @param d: the mapping name to value
    """
    for pv in pvs:
        if pv.Name in d:
            pv.Value = d[pv.Name]


def make_locale(language: str = "", region: str = "",
                subtags: Optional[List[str]] = None) -> UnoStruct:
    """
    Create a Locale object

    @param region: ISO 3166 Country Code.
    @param language: ISO 639 Language Code.
    @param subtags: BCP 47
    @return: the locale (see com.sun.star.lang.Locale)
    """
    locale = uno.createUnoStruct("com.sun.star.lang.Locale")
    if subtags:
        locale.Country = ""
        locale.Language = "qlt"
        if region:
            locale.Variant = "-".join([language, region] + subtags)
        else:
            locale.Variant = "-".join([language] + subtags)
    else:
        locale.Country = region
        locale.Language = language
        locale.Variant = ""
    return locale


def make_border(color: int, width: int,
                style: BorderLineStyle = BorderLineStyle.SOLID) -> UnoStruct:
    """
    Create a border object

    @param color: the color
    @param width: the width
    @param style: the style
    @return: the border (see. com.sun.star.table.BorderLine2)
    """
    border = uno.createUnoStruct("com.sun.star.table.BorderLine2")
    border.Color = color
    border.LineWidth = width
    border.LineStyle = style
    return border


def make_sort_field(field_position: int, asc: bool = True) -> UnoStruct:
    """
    Create a new sort field.

    @param field_position: the position of the field (column number)
    @param asc: True if the sort is ascending, False otherwise
    @return: the sort field (see. com.sun.star.table.TableSortField)
    """
    sf = uno.createUnoStruct("com.sun.star.table.TableSortField")
    sf.Field = field_position
    sf.IsAscending = asc
    return sf


##############################################################################
# RANGES
##############################################################################

def get_last_used_row(oSheet: UnoSheet) -> int:
    """
    Return the last used row

    @param oSheet: the sheet
    @return: the row number (0..)
    """
    return get_used_range_address(oSheet).EndRow


def get_used_range_address(oSheet: UnoSheet) -> UnoRangeAddress:
    """
    Return the used range address (warning: "used" may mean "having data",
    "having formula" or "having some formatting")

    @param oSheet: the sheet
    @return: the used range address (see. com.sun.star.table.CellRangeAddress)
    """
    oCursor = oSheet.createCursor()
    oCursor.gotoStartOfUsedArea(True)
    oCursor.gotoEndOfUsedArea(True)
    return oCursor.RangeAddress


def get_used_range(oSheet: UnoSheet) -> UnoRange:
    """
    Return the used range address (warning: "used" may mean "having data",
    "having formula" or "having some formatting")

    @param oSheet: the sheet
    @return: the used range (see. com.sun.star.table.XCellRange)
    """
    oRangeAddress = get_used_range_address(oSheet)
    return narrow_range_to_address(oSheet, oRangeAddress)


def narrow_range_to_address(
        oSheet: UnoSheet, oRangeAddress: UnoRangeAddress) -> UnoRange:
    """
    Return the range on a sheet at a given cell range address.
    Useful to copy a data array from one sheet to another.

    @param oSheet: the sheet
    @param oRangeAddress: the range address (see. com.sun.star.table.CellRangeAddress)
    @return: the range (see. com.sun.star.table.XCellRange)
    """
    return oSheet.getCellRangeByPosition(
        oRangeAddress.StartColumn, oRangeAddress.StartRow,
        oRangeAddress.EndColumn, oRangeAddress.EndRow)


def get_range_size(oRange: UnoRange) -> Tuple[int, int]:
    """
    Useful for `oRange.getRangeCellByPosition(...)`.

    @param oRange: a SheetCellRange obj
    @return: width and height of the range.
    """
    oAddress = oRange.RangeAddress
    width = oAddress.EndColumn - oAddress.StartColumn + 1
    height = oAddress.EndRow - oAddress.StartRow + 1
    return width, height


def copy_range(oSourceRange: UnoRange):
    """
    Copy a range to the clipboard using the dispatcher.

    @param oSourceRange: the range (may be a sheet)
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
    Paste the copied range to a sheet. Preselected options are the most common:
    * SVDT: strings, values, dates, formats
    * don't skip empty cells, don't transpose, don't copy as link
    * overwrite cells (move mode = 4)

    See core/sc/source/ui/view/cellsh1.cxx
    InsertDeleteFlags FlagsFromString(...) for the list of flags

    @param oDestSheet: the destination sheet
    @param oDestAddress: the destination address (see. com.sun.star.table.CellAddress)
    @param formulas: if True, paste formulas, otherwise paste data (paste special)
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
    Narrow the range to the used range ("used" here means: "having data").

    @param oRange: the range, usually a row or a column
    @param narrow_data: if True, remove top/bottom blank lines and left/right
     blank colmuns
    @return the narrowed range or None if there is nothing left
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
        start_column += left_void_column_count(data_array)
        end_column -= right_void_column_count(data_array)
        oNarrowedRange = oSheet.getCellRangeByPosition(
            start_column, start_row, end_column, end_row)

    return oNarrowedRange


##############################################################################
# DATA ARRAY
##############################################################################
def data_array(oSheet: UnoSheet) -> DATA_ARRAY:
    """
    Return the used data array from a sheet
    @param oSheet: the sheet
    @return: the data array
    """
    return get_used_range(oSheet).DataArray


def is_empty_da_value(v: DATA_VALUE) -> bool:
    """
    Check if a value is empty (void string or spaces)

    @param v: the value in the data array
    @return: True if the value is an empty string or spaces.
    """
    return isinstance(v, str) and v.strip() == ""


def top_void_row_count(data_array: DATA_ARRAY) -> int:
    """
    Count the number of void rows a the top of a a DataArray. A row is void if
    every value is empty (see `is_empty_da_value`).

    Example:
    ```
    top_void_row_count(
        ((" ", " "), ("a", "b"))
    )
    ```
    returns 1

    @param data_array: a data array
    @return: the number of void rows at the top
    """
    r0 = 0
    row_count = len(data_array)
    while r0 < row_count and all(is_empty_da_value(v) for v in data_array[r0]):
        r0 += 1
    return r0


def bottom_void_row_count(data_array: DATA_ARRAY) -> int:
    """
    Count the number of void rows a the bottom of a DataArray. A row is void if
    every value is empty (see `is_empty_da_value`).

    @param data_array: a data array
    @return: the number of void rows at the bottom
    """
    row_count = len(data_array)
    r1 = 0
    # r1 < row_count => row_count - r1 > 0 => row_count - r1 - 1 >= 0
    while r1 < row_count and all(
            is_empty_da_value(v) for v in data_array[row_count - r1 - 1]):
        r1 += 1
    return r1


def left_void_column_count(data_array: DATA_ARRAY) -> int:
    """
    Count the number of void columns a the left of a DataArray. A column is
    void if every value is empty (see `is_empty_da_value`).

    Example:
    ```
    left_void_column_count(
        ((" ", "a"), (" ", "b"))
    )
    ```
    returns 1

    @param data_array: a data array
    @return: the number of void colmuns at the left
    """
    row_count = len(data_array)
    if row_count == 0:
        return 0

    c0 = len(data_array[0])
    for row in data_array:
        c = 0
        while c < c0 and is_empty_da_value(row[c]):
            c += 1
        if c < c0:
            c0 = c
    return c0


def right_void_column_count(data_array: DATA_ARRAY) -> int:
    """
    Count the number of void columns a the right of a DataArray. A column is
    void if every value is empty (see `is_empty_da_value`).

    @param data_array: a data array
    @return: the number of void columns at the right
    """
    row_count = len(data_array)
    if row_count == 0:
        return 0

    width = len(data_array[0])
    c1 = width
    for row in data_array:
        c = 0
        while c < c1 and is_empty_da_value(row[width - c - 1]):
            c += 1
        if c < c1:
            c1 = c
    return c1


###############################################################################
# FORMATTING
###############################################################################

class ListValidationBuilder:
    """
    A builder for a cell validation based on a list of values.
    """

    def __init__(self):
        self._values = []
        self._default_string = None
        self._ignore_blank = False
        self._sort_values = False
        self._show_error = True

    def values(self, values: List[Any]) -> "ListValidationBuilder":
        """
        Set the allowed values

        @param values: the values
        @return: self
        """
        self._values = values
        return self

    def default_string(self, default_string: str) -> "ListValidationBuilder":
        """
        Does nothing

        @param default_string: a string
        @return: self
        @deprecated: do not use
        """
        self._default_string = default_string
        return self

    def ignore_blank(self, ignore_blank: bool) -> "ListValidationBuilder":
        """
        Allow a blank cell

        @param ignore_blank: if True, blank cell is allowed
        @return: self
        """
        self._ignore_blank = ignore_blank
        return self

    def sort_values(self, sorted_values: bool) -> "ListValidationBuilder":
        """
        Sort the values

        @param sorted_values: if True, sort the available values
        @return: self
        """
        self._sort_values = sorted_values
        return self

    def show_error(self, show_error: bool) -> "ListValidationBuilder":
        """
        Enable the error message.

        @param show_error: if True, show an error message for incorrect values
        @return: self
        """
        self._show_error = show_error
        return self

    def update(self, oCell: UnoCell):
        """
        Update the cell with the current validation.
        @param oCell: the cell
        """
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


class ValidationFactory:
    """
    A factory to create a cell validation.
    """

    def list(self) -> ListValidationBuilder:
        """
        The only method available: allow only values or strings specified in a
        list.
        @return:
        """
        return ListValidationBuilder()


def set_validation_list_by_cell(
        oCell: UnoCell, values: List[Any],
        default_string: Optional[str] = None,
        ignore_blank: bool = True, sorted_values: bool = False,
        show_error: bool = True):
    """
    Add a validation list to a cell.

    @param oCell: the cell
    @param values: the available values
    @param default_string: the default string
    @param ignore_blank: if True, blank cell is allowed
    @param sorted_values: if True, sort the values
    @param show_error: if True, show the error message
    """
    factory = ValidationFactory().list().values(values)
    factory.ignore_blank(ignore_blank)
    factory.sort_values(sorted_values)
    factory.show_error(show_error)
    factory.update(oCell)

    if default_string is not None:
        oCell.String = default_string


def quote_element(value: Any) -> str:
    """
    Quote a list validation element. List validation values
    can only be strings. This function convert all types of values to
    the string type.

    TODO: use a locale

    @param value: the value
    @return: the quoted value
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
    Sort a given range.

    Example:
    ```
    oRange = ...
    sf1 = make_sort_field(0) # column A, asc
    sf1.IsCaseSensitive = True
    sf2 = make_sort_field(4, False) # column E, desc
    sort_range(oRange, (sf1, sf2), True)
    ```

    @param oRange: the range to sort
    @param sort_fields: the sort fields (see: com.sun.star.table.TableSortField)
    @param has_header: True if the range has a header
    """
    typed_sort_fields = uno.Any('[]com.sun.star.table.TableSortField',
                                sort_fields)
    sort_descriptor = oRange.createSortDescriptor()
    update_pvs(sort_descriptor,
               {'ContainsHeader': has_header, 'SortFields': typed_sort_fields})
    oRange.sort(sort_descriptor)


# CONDITIONAL
def clear_conditional_format(oColoredColumns: UnoRange):
    """
    Remove the conditional format

    @param oColoredColumns: the range having a conditional format
    """
    oConditionalFormat = oColoredColumns.ConditionalFormat
    oConditionalFormat.clear()


def conditional_format_on_formulas(
        oColoredColumns: UnoRange, style_by_formula: Mapping[str, str],
        oSrcAddress: UnoCellAddress):
    """
    Add a list of conditional formats to a cell. Each conditionnal format
    is given by a style name and a formula.

    Example:
    ```
    oCell = ...
    conditional_format_on_formulas(oCell, {
        "red_style": "A1 < 0",
        "italic_style": "A1 > 100",
    }, oCell.RangeAddress)
    ```

    @param oColoredColumns: the target columns
    @param style_by_formula: a mapping formula -> style name
    @param oSrcAddress: the common source position for the relative reference
    cells in formulas
    """
    oConditionalFormat = oColoredColumns.ConditionalFormat
    for formula, style in style_by_formula.items():
        oConditionalEntry = get_formula_conditional_entry_values(
            formula, style, oSrcAddress)
        oConditionalFormat.addNew(oConditionalEntry)

    oColoredColumns.ConditionalFormat = oConditionalFormat


def get_formula_conditional_entry_values(
        formula: str, style_name: str, oSrcAddress: UnoCellAddress
) -> Tuple[UnoPropertyValue, ...]:
    """
    Return a condition format entry. It seems surpring, but we don't build
    a com.sun.star.sheet.TableConditionalEntry, instead a tuple of property
    values. See. com.sun.star.sheet.XSheetConditionalEntries.addNew(...)

    @param formula: the value or formula for the operation.
    @param style_name: the name of the style to apply
    @param oSrcAddress: the base address for relative cell references in formulas.
    @return: a tuple of property values
    """
    return get_conditional_entry_values(formula, "", ConditionOperator.FORMULA,
                                        style_name, oSrcAddress)


def get_conditional_entry_values(
        formula1: str, formula2: str, operator: str, style_name: str,
        oSrcAddress: UnoCellAddress) -> Tuple[UnoPropertyValue, ...]:
    """
    Return a condition format entry. It seems surpring, but we don't build
    a com.sun.star.sheet.TableConditionalEntry, instead a tuple of property
    values. See. com.sun.star.sheet.XSheetConditionalEntries.addNew(...)

    @param formula1: the value or formula for the operation.
    @param formula2: the second value or formula for the operation
    (used with ConditionOperator::BETWEEN or ConditionOperator::NOT_BETWEEN operations).
    @param operator: the operation to perform for this condition.
    @param style_name: the name of the style to apply
    @param oSrcAddress: the base address for relative cell references in formulas.
    @return: a tuple of property values
    """
    return make_pvs({
        "Formula1": formula1, "Formula2": formula2, "Operator": operator,
        "StyleName": style_name, "SourcePosition": oSrcAddress
    })


def find_or_create_number_format_style(oDoc: UnoSpreadsheetDocument, fmt: str,
                                       locale: Optional[UnoStruct] = None
                                       ) -> int:
    """
    Get the key for a number format. If the format does not exist, create it.

    @param oDoc: the document
    @param fmt: the format string
    @param locale: the locale (if None, the current locale is used)
    @return: the format key
    """
    # from Andrew Pitonyak 5.14
    # www.openoffice.org/documentation/HOW_TO/various_topics/AndrewMacro.odt
    oFormats = oDoc.NumberFormats
    if locale is None:
        oLocale = make_locale()
    else:
        oLocale = locale
    format_key = oFormats.queryKey(fmt, oLocale, True)
    if format_key == -1:
        format_key = oFormats.addNew(fmt, oLocale)

    return format_key


def create_filter(oRange: UnoRange):
    """
    Create a new auto filter.

    @param oRange: the range to filter
    """
    oDoc = parent_doc(oRange)
    oController = oDoc.CurrentController
    oController.select(oRange)
    provider.dispatcher.executeDispatch(
        oController, ".uno:DataFilterAutoFilter", "", 0, tuple())
    # unselect
    oRanges = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
    oController.select(oRanges)


def remove_filter(oRange: UnoRange):
    """
    Remove the existing filter on a range.

    @param oRange: The range
    """
    # True means "empty"
    oFilterDescriptor = oRange.createFilterDescriptor(True)
    oRange.filter(oFilterDescriptor)
    create_filter(oRange)


def row_as_header(oHeaderRow: UnoRow):
    """
    Format a row as the header.

    @param oHeaderRow: the row (usually the first row)
    """
    oHeaderRow.CharWeight = FontWeight.BOLD
    oHeaderRow.CharWeightAsian = FontWeight.BOLD
    oHeaderRow.CharWeightComplex = FontWeight.BOLD
    oHeaderRow.IsTextWrapped = True
    oHeaderRow.OptimalHeight = True


def column_optimal_width(oColumn: UnoColumn, min_width: int = 2 * 1000,
                         max_width: int = 10 * 1000):
    """
    Set the width of the column to an optimal value.

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
    Return the page style.

    @param oSheet: a sheet
    @return: the page style of this sheet (see. com.sun.star.style.PageStyle)
    """
    page_style_name = oSheet.PageStyle
    oDoc = parent_doc(oSheet)
    oStyle = oDoc.StyleFamilies.getByName("PageStyles").getByName(
        page_style_name)
    return oStyle  # type: ignore[return-value]


def set_paper(oSheet: UnoSheet):
    """
    Set the paper size for this sheet.

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
        text_field_url = cast(UnoService, oDoc.createInstance(
            "com.sun.star.text.TextField.URL"))
        text_field_url.Representation = line
        text_field_url.URL = url
        oCell.insertTextContent(oCursor, text_field_url, False)


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
    """List of new document URLs"""
    Calc = "private:factory/scalc"
    Writer = "private:factory/swriter"
    Draw = "private:factory/sdraw"
    Impress = "private:factory/simpress"
    Math = "private:factory/smath"


class Target(str, Enum):
    """List of new document targets"""
    BLANK = "_blank"  # always creates a new frame
    # special UI functionality (e.g. detecting of already loaded documents,
    # using of empty frames of creating of new top frames as fallback)
    DEFAULT = "_default"
    SELF = "_self"  # means frame himself
    PARENT = "_parent"  # address direct parent of frame
    TOP = "_top"  # indicates top frame of current path in tree
    BEAMER = "_beamer"  # means special sub frame


def open_document(filename: Union[str, Path], target: str = Target.BLANK,
                  frame_flags=FrameSearchFlag.AUTO,
                  **kwargs) -> UnoSpreadsheetDocument:
    """
    Open a document in LibreOffice.

    Example:
    ```
    oDoc = open_document("a_text_doc.odt", Hidden=True)
    ```

    Other example:
    ```
    oDoc = open_document("a_csv_doc.csv",
        FilterName=Filter.CSV,
        FilterOptions=create_import_filter_options(delimiter=';', encoding='latin-1'),
        Hidden=True
    )
    ```
    (You can also open a void calc document and then use `import_csv`)

    @param filename: the name of the file to open
    @param target: "the name of the frame to view the document in" or a special
    target
    @param frame_flags: where to search the frame
    @param kwargs: the parameters, see com.sun.star.document.MediaDescriptor.
    For convert filters, see.
    https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html
    For CSV filter options, see.
    https://help.libreoffice.org/latest/en-US/text/shared/guide/csv_params.html
    @return: a reference on the doc
    """
    url = uno_path_to_url(filename)
    if kwargs:
        params = make_pvs(kwargs)
    else:
        params = ()
    return cast(
        UnoSpreadsheetDocument,
        provider.desktop.loadComponentFromURL(url, target, frame_flags, params)
    )


"""@deprecated use open_document"""
open_in_calc = open_document


# Create a document
####
def doc_builder(
        url: NewDocumentUrl = NewDocumentUrl.Calc,
        taget_frame_name: Target = Target.BLANK,
        search_flags: FrameSearchFlag = FrameSearchFlag.AUTO,
        pvs: Optional[UnoPropertyValues] = None
) -> "DocBuilder":
    """
    Create a new document builder. See.
    com.sun.star.frame.XComponentLoader.loadComponentFromURL.

    @param url: the url (a NewDocumentUrl)
    @param taget_frame_name: the target frame
    @param search_flags: the search flags (see. FrameSearchFlags)
    @param pvs: the property values
    @return: a DocBuilder instance
    """
    if pvs is None:
        pvs = tuple()
    return DocBuilder(url, taget_frame_name, search_flags, pvs)


def new_doc(url: NewDocumentUrl = NewDocumentUrl.Calc,
            taget_frame_name: Target = Target.BLANK,
            search_flags: FrameSearchFlag = FrameSearchFlag.AUTO,
            pvs: Optional[UnoPropertyValues] = None) -> UnoSpreadsheetDocument:
    """
    Create a blank new doc

    @param url: the url (a NewDocumentUrl)
    @param taget_frame_name: the target frame
    @param search_flags: the search flags (see. FrameSearchFlags)
    @param pvs: the property values
    @return: the doc (see com.sun.star.lang.XComponent)
    """
    return doc_builder(url, taget_frame_name, search_flags, pvs).build()


class DocBuilder:
    """
    A document Calc builder.

    Example:
    ```
    d = doc_builder(NewDocumentUrl.Calc)
    d.sheet_names(["main", "aux"]
                  trunc_if_necessary=False)
    oDoc = d.build()
    ```
    """

    def __init__(
            self, url: NewDocumentUrl, target_frame_name: Target,
            search_flags: FrameSearchFlag,
            pvs: UnoPropertyValues):
        """Create a blank new doc"""
        self._url = url
        self._target_frame_name = target_frame_name
        self._search_flags = search_flags
        self._pvs = pvs
        self._sheet_names = cast(List[str], None)
        self._expand_if_necessary = True
        self._trunc_if_necessary = True
        self._apply_func = cast(Callable[[UnoSheet], None], None)
        self._apply_funcs = cast(List[Callable[[UnoSheet], None]], None)
        self._make_func = cast(Callable[[UnoSheet], None], None)
        self._duplicate_sheet_names = cast(List[str], None)
        self._remove = True
        self._duplicate_to = cast(int, None)
        self._final_sheet_count = cast(int, None)

    def build(self) -> UnoSpreadsheetDocument:
        oDoc = cast(
            UnoSpreadsheetDocument,
            provider.desktop.loadComponentFromURL(
                self._url, self._target_frame_name, self._search_flags,
                self._pvs)
        )
        oDoc.lockControllers()
        try:
            self._build_sheet_names(oDoc)
            self._build_apply_func_to_sheets(oDoc)
            self._build_apply_func_list_to_sheets(oDoc)
            self._build_make_base_sheet(oDoc)
            self._build_duplicate_base_sheet(oDoc)
            self._build_duplicate_to(oDoc)
            self._build_trunc_to_count(oDoc)
        finally:
            oDoc.unlockControllers()
        return oDoc

    def _build_sheet_names(self, oDoc: UnoSpreadsheetDocument):
        if self._sheet_names is None:
            return

        oSheets = oDoc.Sheets
        it = iter(self._sheet_names)
        s = 0

        initial_count = oSheets.Count
        try:
            # rename
            while s < initial_count:
                oSheet = oSheets.getByIndex(s)
                oSheet.Name = next(it)  # may raise a StopIteration
                s += 1

            if s != initial_count:
                raise AssertionError("s={} vs oSheets.Count={}".format(
                    s, initial_count))

            if self._expand_if_necessary:
                # add
                for sheet_name in it:
                    oSheets.insertNewByName(sheet_name, s)
                    s += 1
        except StopIteration:  # it
            if s > initial_count:
                raise AssertionError("s={} vs oSheets.Count={}".format(
                    s, oSheets.getCount()))
            if self._trunc_if_necessary:
                self._trunc_to_count(oDoc, s)

        return self

    def _build_apply_func_to_sheets(
            self, oDoc: UnoSpreadsheetDocument):
        if self._apply_func is None:
            return

        oSheets = oDoc.Sheets
        for oSheet in to_iter(oSheets):
            self._apply_func(oSheet)

    def _build_apply_func_list_to_sheets(
            self, oDoc: UnoSpreadsheetDocument):
        if self._apply_funcs is None:
            return

        oSheets = oDoc.Sheets
        for func, oSheet in zip(self._apply_funcs, to_iter(oSheets)):
            func(oSheet)
        return self

    def _build_make_base_sheet(self, oDoc: UnoSpreadsheetDocument):
        if self._make_func is None:
            return

        oSheets = oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        self._make_func(oBaseSheet)

    def _build_duplicate_base_sheet(self, oDoc: UnoSpreadsheetDocument):
        if self._duplicate_sheet_names is None:
            return

        oSheets = oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        self._make_func(oBaseSheet)
        for s, sheet_name in enumerate(self._duplicate_sheet_names, 1):
            oSheets.copyByName(oBaseSheet.Name, sheet_name, s)

        if self._remove:
            while oSheets.Count > len(self._duplicate_sheet_names):
                oSheets.removeByName(oSheets.getByIndex(oSheets.Count - 1).Name)

    def _build_duplicate_to(self, oDoc: UnoSpreadsheetDocument):
        if self._duplicate_to is None:
            return

        oSheets = oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        for s in range(self._duplicate_to + 1):
            oSheets.copyByName(oBaseSheet.Name, oBaseSheet.Name + str(s), s)

    def _build_trunc_to_count(self, oDoc: UnoSpreadsheetDocument):
        if self._final_sheet_count is None:
            return

        self._trunc_to_count(oDoc, self._final_sheet_count)

    def _trunc_to_count(self, oDoc: UnoSpreadsheetDocument,
                        final_sheet_count: int):
        oSheets = oDoc.Sheets
        while oSheets.Count > final_sheet_count:
            oSheet = oSheets.getByIndex(final_sheet_count)
            oSheets.removeByName(oSheet.Name)

    # methods
    def sheet_names(self, sheet_names: List[str],
                    expand_if_necessary: bool = True,
                    trunc_if_necessary: bool = True) -> "DocBuilder":
        """
        Set the document sheet names

        @param sheet_names: the sheet names
        @param expand_if_necessary: if there are more sheet names than sheets,
        create new sheets for remaining sheet names
        @param trunc_if_necessary: if there are fewer sheet names than sheets,
        remove the extra sheets
        @return: self
        """
        self._sheet_names = sheet_names
        self._expand_if_necessary = expand_if_necessary
        self._trunc_if_necessary = trunc_if_necessary
        return self

    def apply_func_to_sheets(
            self, func: Callable[[UnoSheet], None]) -> "DocBuilder":
        """
        Apply the same function to every sheet of the document

        @param func: the function
        @return: self
        """
        self._apply_func = func
        return self

    def apply_func_list_to_sheets(
            self, funcs: List[Callable[[UnoSheet], None]]) -> "DocBuilder":
        """
        Apply a different function to each sheet of the document.
        If there are fewer functions than sheets, the last sheets are ignored.
        If there are more functions than sheets, the last functions are ignored.

        @param funcs: the functions
        @return: self
        """
        self._apply_funcs = funcs
        return self

    def duplicate_base_sheet(self, func: Callable[[UnoSheet], None],
                             sheet_names: List[str], remove: bool = True):
        """
        Create a base sheet and duplicate it n-1 times

        @param func: the function to create the base sheet
        @param sheet_names: the names of the duplicated sheets
        @param remove: if True, remove other sheets
        @return: self
        """
        self._make_func = func
        self._duplicate_sheet_names = sheet_names
        self._remove = remove
        return self

    def make_base_sheet(self, func: Callable[[UnoSheet], None]
                        ) -> "DocBuilder":
        """
        Create a base sheet.

        @param func: the function to create the base sheet
        @return: self
        """
        self._make_func = func
        return self

    def duplicate_to(self, n: int) -> "DocBuilder":
        """
        Duplicate the first sheet n times

        @param n: the number of the last sheet
        @return: self
        """
        self._duplicate_to = n
        return self

    def trunc_to_count(self, final_sheet_count: int) -> "DocBuilder":
        """
        Remove all sheets after the sheet number final_sheet_count - 1.

        @param final_sheet_count: the maximum final number of sheets
        @return: self
        """
        self._final_sheet_count = final_sheet_count
        return self


###############################################################################
# MISC
###############################################################################
_ApplyType = Callable[
    [str, Union[DATA_ROW, DATA_VALUE]], Tuple[str, Union[DATA_ROW, DATA_VALUE]]
]


def read_options(
        oSheet: UnoSpreadsheetDocument, aAddress: UnoRangeAddress,
        apply: Optional[_ApplyType] = None
) -> Mapping[str, Union[DATA_ROW, DATA_VALUE]]:
    """
    Read options stored in a sheet.

    @param oSheet: the "Options" sheet
    @param aAddress: the options range address
    @param apply: an optional function to apply to (key, values) pairs
    @return: a mapping key -> value or list of values
    """
    options = {}
    if aAddress.StartColumn == aAddress.EndColumn:
        return {}

    oRange = oSheet.getCellRangeByPosition(
        aAddress.StartColumn, aAddress.StartRow,
        aAddress.EndColumn, aAddress.EndRow)
    data_array = oRange.DataArray
    for row in data_array:
        k = str(row[0])
        v = rtrim_row(row[1:])
        if apply is None:
            new_k, new_v = k, v
        else:
            new_k, new_v = apply(k, v)
        if new_k:
            options[new_k] = new_v
    return options


def rtrim_row(row: DATA_ROW, null="") -> Union[DATA_ROW, DATA_VALUE]:
    """
    Given a list of values, remove the empty values.
    Then:
        * If the list is empty, return null.
        * If the list has one element, return the element.
        * Else return the list.

    @param row: the list of values
    @param null: the null value
    @return: a value or a list of values
    """
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
        apply: Optional[_ApplyType]
) -> Mapping[str, Union[DATA_ROW, DATA_VALUE]]:
    """
    Read options stored in a sheet.

    @param oDoc: the document
    @param sheet_name: the name of the sheet
    @param apply: an optional function to apply to (key, values) pairs
    @return: a mapping key -> value
    """
    oSheet = oDoc.Sheets.getByName(sheet_name)
    oRangeAddress = get_used_range_address(oSheet)
    return read_options(oSheet, oRangeAddress, apply)


def copy_row_at_index(oSheet: UnoSheet, row: DATA_ROW, r: int):
    """
    Copy the row (as a tuple of values) to a given index.

    @param oSheet: the sheet
    @param row: the data row
    @param r: the index
    """
    oRange = oSheet.getCellRangeByPosition(0, r, len(row) - 1, r)
    oRange.DataArray = [row]


class _Inspector:
    """
    An inspector: MRI or Xray provider.
    """

    def __init__(self, provider: _ObjectProvider):
        """
        @param provider: the object provider
        """
        self._provider = provider
        self._xray_script = lazy(UnoService)
        self._ignore_xray = False
        self._oMRI = lazy(UnoService)
        self._ignore_mri = False

    def use_xray(self, fail_on_error: bool = False):
        """
        Try to load Xray lib.
        @param fail_on_error: Should this function fail on error
        @raises UnoRuntimeException: if Xray is not avaliable and
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

        @param obj: the obj
        @param fail_on_error: Should this function fail on error
        @raises RuntimeException: if Xray is not avaliable and `fail_on_error`
         is True.
        """
        if self._ignore_xray:
            return

        if self._xray_script is None:
            self.use_xray(fail_on_error)
            if self._ignore_xray:
                return

        _oi = cast(Tuple[int, ...], tuple())
        _o = cast(Tuple[Any, ...], tuple())
        self._xray_script.invoke((obj,), _oi, _o)

    def mri(self, obj: Any, fail_on_error: bool = False):
        """
        MRI an object.

        @param fail_on_error: if True and MRI is not installed, raise an exception
        @param obj: the object
        @raise UnoRuntimeException: if fail_on_error is True and MRI is not
        installed
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
                else:
                    return

        self._oMRI.inspect(obj)


def get_inspector() -> _Inspector:
    return _inspector


TEXT_FLAVOR = ("text/plain;charset=utf-16", "Unicode-text")
HTML_FLAVOR = ("text/html;charset=utf-8", "HTML")


def copy_to_clipboard(value: Any, flavor: Tuple[str, str] = TEXT_FLAVOR):
    """
    Copy the value to the clipboard.

    See https://forum.openoffice.org/en/forum/viewtopic.php?t=93562

    @param value: the value
    @param flavor: the flavor as tuple (mimetype name, human readable name)
    """
    oClipboard = create_uno_service(
        "com.sun.star.datatransfer.clipboard.SystemClipboard")
    oClipboard.setContents(Transferable(value, flavor), None)


def get_from_clipboard(flavor: Tuple[str, str] = TEXT_FLAVOR) -> Optional[Any]:
    """
    Retrieves the value from the clipboard.

    See https://forum.openoffice.org/en/forum/viewtopic.php?t=93562

    @param flavor: the flavor as tuple (mimetype name, human readable name)
    @return: the content of the clipboard or None
    """
    oClipboard = create_uno_service(
        "com.sun.star.datatransfer.clipboard.SystemClipboard")
    oContents = oClipboard.getContents()
    oTypes = oContents.getTransferDataFlavors()

    for oType in oTypes:
        if oType.MimeType == flavor[0]:
            return oContents.getTransferData(oType)

    return None


class Transferable(unohelper.Base, XTransferable):
    """
    A com.sun.star.datatransfer.XTransferable object, carrying a value
    for a flavor.
    """

    def __init__(self, value: Any, flavor: Tuple[str, str]):
        """
        @param value: the value
        @param flavor: the flavor as tuple (mimetype name, human readable name)
        """
        self._value = value
        self._flavor = flavor

    def getTransferData(self, aFlavor: UnoObject):
        """see. com.sun.star.datatransfer.XTransferable.getTransferData"""
        if aFlavor.MimeType == self._flavor[0]:
            return self._value

    def getTransferDataFlavors(self):
        """see. com.sun.star.datatransfer.XTransferable.getTransferDataFlavors"""
        flavor = create_uno_struct("com.sun.star.datatransfer.DataFlavor",
                                   MimeType=self._flavor[0],
                                   HumanPresentableName=self._flavor[1])
        return [flavor]

    def isDataFlavorSupported(self, aFlavor: UnoObject) -> bool:
        """see. com.sun.star.datatransfer.XTransferable.isDataFlavorSupported"""
        return aFlavor.MimeType == self._flavor[0]


def convert_to_html(oTextRange: UnoTextRange) -> str:
    """
    Convert a sequence of chars in a text range to HTML

    @param oTextRange: the text range
    @return: the HTML string
    """
    return HTMLConverter().convert(oTextRange)


class HTMLConverter:
    """
    Minimalist chars to HTML converter
    """
    _logger = logging.getLogger(__name__)

    def __init__(self, html_line_break: str = "<br>"):
        self._html_line_break = html_line_break

    def convert(self, oTextRange: UnoTextRange) -> str:
        """
        Convert a sequence of chars in a text range to HTML

        @param oTextRange: the text range
        @return: the HTML string
        """
        html = self._html_line_break.join(
            self._par_to_html(par_text_range) for par_text_range in
            to_iter(oTextRange))
        return html

    def _par_to_html(self, par_text_range: UnoTextRange) -> str:
        return "".join(
            [self._to_html(chunk) for chunk in to_iter(par_text_range)])

    def _to_html(self, text_range: UnoTextRange) -> str:
        tag = self._get_tag(text_range)

        statements = []
        if text_range.CharFontName != "Liberation Sans":
            statements.append(
                "font-family: \"{}\"".format(text_range.CharFontName))
        if text_range.CharHeight != 10:
            statements.append("font-size: {}pt".format(text_range.CharHeight))
        if text_range.CharWeight != 100:
            statements.append(
                "font-weight: {}".format(int(text_range.CharWeight * 4)))
        italic = (text_range.CharPosture == FontSlant.OBLIQUE
                  or text_range.CharPosture == FontSlant.ITALIC)
        if italic:
            statements.append("font-style: italic")
        if text_range.CharColor != -1:
            statements.append("color: #{:02x}".format(text_range.CharColor))
        if text_range.CharBackColor != -1:
            statements.append(
                "background-color: #{:02x}".format(text_range.CharBackColor))
        if text_range.CharOverline != 0:
            statements.append("text-decoration: overline")
        if text_range.CharStrikeout != 0:
            statements.append("text-decoration: line-through")
        if text_range.CharUnderline != 0:
            statements.append("text-decoration: underline")

        if (text_range.TextPortionType == "TextField"
                and text_range.TextField.supportsService(
                    "com.sun.star.text.TextField.URL")
        ):
            text = "<a href='{}'>{}</a>".format(
                text_range.TextField.URL, text_range.TextField.Representation)
        else:
            text = text_range.String
        if statements:
            return "<{tag} style='{style}'>{text}</{tag}>".format(
                tag=tag, style="; ".join(statements), text=text
            )
        elif tag == "span":
            return text
        else:
            return "<{tag}>{text}</{tag}>".format(
                tag=tag, text=text
            )

    def _get_tag(self, text_range: UnoTextRange) -> str:
        if text_range.CharEscapementHeight < 100:
            if text_range.CharEscapement < 0:
                return "sub"
            elif text_range.CharEscapement > 0:
                return "sup"
        return "span"


class SheetFormatter:
    """
    A helper to format a sheet
    """

    def __init__(self, oSheet: UnoSheet, locale: Optional[UnoStruct] = None):
        """
        @param oSheet: the sheet to format
        @param locale: the local, or None.
        """
        self._oSheet = oSheet
        oDoc = parent_doc(oSheet)
        self._oFormats = oDoc.NumberFormats
        if locale is None:
            self._locale = self._find_locale()
        else:
            self._locale = locale

    @staticmethod
    def _find_locale() -> UnoStruct:
        """
        @return: the current locale as a UNO structure
        """
        try:
            language_code = cast(str, getlocale()[0])
            region, language = language_code.split("_")
            return make_locale(language, region)
        except (IndexError, ValueError):
            return make_locale("EN", "us")

    def first_row_as_header(self):
        """
        Set the first row a a header (see: row_as_header)
        """
        oRows = self._oSheet.Rows
        row_as_header(oRows.getByIndex(0))

    def set_format(self, fmt_name: str, *col_indices: int):
        """
        Set a format to a list of columns

        @param fmt_name: the name of the format
        @param col_indices: the column indices
        """
        number_format_id = self._oFormats.queryKey(
            fmt_name, self._locale, True)
        if number_format_id == -1:
            number_format_id = self._oFormats.addNew(fmt_name, self._locale)

        oColumns = self._oSheet.Columns
        for i in col_indices:
            oColumns.getByIndex(i).NumberFormat = number_format_id

    def fix_first_row(self):
        """
        Fix the first row
        """
        set_print_area(self._oSheet, self._oSheet.Rows.getByIndex(0))

    def set_print_area(self):
        """
        Set the print area (see. set_print_area)
        """
        set_print_area(self._oSheet)

    def set_optimal_width(self, *col_indices: int, min_width: int = 2 * 1000,
                          max_width: int = 10 * 1000):
        """
        Set the optimal width for some columns (see. column_optimal_width).

        @param col_indices: the column indices
        @param min_width: the min width
        @param max_width: the max width
        @return:
        """
        oColumns = self._oSheet.Columns
        for i in col_indices:
            column_optimal_width(oColumns.getByIndex(i), min_width, max_width)

    def create_filter(self):
        """
        Create an auto filter (see. create_filter)
        """
        create_filter(self._oSheet)


class DatesHelper:
    """
    A helper for dates, using the NullDate of a document (see
    com.sun.star.sheet.SpreadsheetDocumentSettings.NullDate).

    Example:

    ```
    helper = DatesHelper.create(oDoc)
    helper.date_to_int(dt.date(1899, 12, 30)) # 0 if this is the NullDate
    ```
    """

    @staticmethod
    def create(oDoc: Optional[UnoSpreadsheetDocument] = None) -> "DatesHelper":
        """
        @param oDoc: the document that has a NullDate
        @return: the helper
        """
        if oDoc is None:
            origin = dt.datetime(1899, 12, 30)
        else:
            oNullDate = oDoc.NullDate
            if oNullDate.Day == 0:
                origin = dt.datetime(1899, 12, 30)
            else:
                origin = dt.datetime(oNullDate.Year, oNullDate.Month,
                                     oNullDate.Day)
        return DatesHelper(origin)

    def __init__(self, origin: dt.datetime):
        """
        @param origin: the NullDate as Python datetime.
        """
        self._origin = origin

    def date_to_int(self, a_date: Union[dt.date, dt.datetime]) -> int:
        """
        Converts a date to an int.

        @param a_date: the Python date or datetime to convert
        @return: the LibreOffice internal representation of the date as an int.
        """
        if isinstance(a_date, dt.datetime):
            pass
        elif isinstance(a_date, dt.date):
            a_date = dt.datetime(a_date.year, a_date.month, a_date.day)
        else:
            raise ValueError()
        return (a_date - self._origin).days

    def date_to_float(self,
                      a_date: Union[dt.date, dt.datetime, dt.time]) -> float:
        """
        Converts a date to a float.

        @param a_date: the Python date or datetime to convert
        @return: the LibreOffice internal representation of the date as a float.
        """
        if isinstance(a_date, dt.datetime):
            time_delta = a_date - self._origin
        elif isinstance(a_date, dt.date):
            a_datetime = dt.datetime(a_date.year, a_date.month, a_date.day)
            time_delta = a_datetime - self._origin
        elif isinstance(a_date, dt.time):
            time_delta = dt.timedelta(
                hours=a_date.hour, minutes=a_date.minute, seconds=a_date.second,
                microseconds=a_date.microsecond)
        else:
            raise ValueError(a_date)
        return time_delta.total_seconds() / 86400

    def int_to_date(self, days: int) -> dt.datetime:
        """
        Converts an int to a date.

        @param days: the LibreOffice internal representation of the date as an int
        @return: the date
        """
        return self._origin + dt.timedelta(days)

    def float_to_date(self, days: float) -> dt.datetime:
        """
        Converts a float to a date.

        @param days: the LibreOffice internal representation of the date as a float
        @return: the datetime
        """
        return self._origin + dt.timedelta(days)


def date_to_int(a_date: Union[dt.date, dt.datetime]) -> int:
    """
    Converts a date to an int, using default origin (1899-12-30)

    @deprecated: use DateHelper
    @param a_date: the Python date or datetime to convert
    @return: the LibreOffice internal representation of the date as an int.
    """
    return DatesHelper.create().date_to_int(a_date)


def date_to_float(a_date: Union[dt.date, dt.datetime, dt.time]) -> float:
    """
    Converts a date to a float, using default origin (1899-12-30)

    @deprecated: use DateHelper
    @param a_date: the Python date or datetime to convert
    @return: the LibreOffice internal representation of the date as a float.
    """
    return DatesHelper.create().date_to_float(a_date)


def int_to_date(days: int) -> dt.datetime:
    """
    Converts an int to a date, using default origin (1899-12-30)

    @deprecated: use DateHelper
    @param days: the LibreOffice internal representation of the date as an int.
    @return: the Python datetime
    """
    return DatesHelper.create().int_to_date(days)


def float_to_date(days: float) -> dt.datetime:
    """
    Converts a float to a date, using default origin (1899-12-30)

    @deprecated: use DateHelper
    @param days: the LibreOffice internal representation of the date as a float.
    @return: the Python datetime
    """
    return DatesHelper.create().float_to_date(days)


def copy_data_array(
        oCell: UnoCell,
        data_array: DATA_ARRAY, undo=True, debug=False, step=10000,
        callback: Callable[[int], None] = None):
    """
    Copy a data array to a given address on a sheet. This function provides
    some convenient helpers.
    If `undo` is False, don't add the action on the undo stack. Since data array
    can be a memory expensive operation, storing it on the stack may use a lot
    of memory.
    If `debug` is True, the function will check if the data array is a rectangle
    and if all values are authorized. If the check fails, a ValueError is
    raised.
    The `chunk_size` will set the size in rows of the slices to be copied. This helps
    copying without using to much memory. If `chunk_size` is -1, then the copy
    operation is done all at once.
    The `callback` function is called after every chunk_size, with a parameter:
    the number of rows processed. The `callback` is always called once at
    the end of the copy.

    @param oCell: the top cell of the destination array
    @param data_array: the data array
    @param undo: if False, don't add to undo stack.
    @param debug: if True, check the data array
    @param step: size of the slice in rows
    @param callback: function called after every chunk_size
    """
    DataArrayCopier(undo, debug, step, callback).copy(oCell, data_array)


class DataArrayCopier:
    """
    A copier for data arrays.

    If `undo` is False, don't add the action on the undo stack. Since data array
    can be a memory expensive operation, storing it on the stack may use a lot
    of memory.
    If `debug` is True, the function will check if the data array is a rectangle
    and if all values are authorized. If the check fails, a ValueError is
    raised.
    The `chunk_size` will set the size in rows of the slices to be copied. This helps
    copying without using to much memory. If `chunk_size` is -1, then the copy
    operation is done all at once.
    The `callback` function is called after every chunk_size, with a parameter:
    the number of rows processed. The `callback` is always called once at
    the end of the copy.

    @param undo: if False, don't add to undo stack.
    @param debug: if True, check the data array
    @param step: size of the slice in rows
    @param callback: function called after every chunk_size
    """

    def __init__(self, undo=True, debug=False, chunk_size=10000,
                 callback: Callable[[int], None] = None):
        self._undo = undo
        self._debug = debug
        self._chunk_size = chunk_size
        if callback is None:
            def _ignore_callback(_start_r: int):
                pass

            self._callback = _ignore_callback
        else:
            self._callback = callback

    def copy(self, oCell: UnoCell, data_array: DATA_ARRAY):
        """
        @param oDoc: the document
        @param cell_address: the cell address of the top left cell of the array
        @param data_array: the data array
        """
        if self._debug:
            self.check_data_array(data_array)

        row_count = len(data_array)
        if row_count == 0:
            return

        col_count = len(data_array[0])
        if col_count == 0:
            return

        oDoc = parent_doc(oCell)
        if self._undo:
            if self._chunk_size < 0 or row_count <= self._chunk_size:
                self._copy_whole_data_array(oCell, data_array)
            else:
                with undo_context(oDoc):
                    self._copy_data_array_by_chunks(oCell, data_array)
        else:
            if self._chunk_size < 0 or row_count <= self._chunk_size:
                with no_undo_context(oDoc):
                    self._copy_whole_data_array(oCell, data_array)
            else:
                with no_undo_context(oDoc):
                    self._copy_data_array_by_chunks(oCell, data_array)

    @staticmethod
    def check_data_array(data_array: DATA_ARRAY):
        """
        Check the data_array format.
        @param data_array:
        @return:
        """
        row_count = len(data_array)
        if row_count == 0:
            return

        col_count = len(data_array[0])

        square_errs = []
        illegal_value_errs = []
        for i, row in enumerate(data_array):
            if len(row) != col_count:
                square_errs.append(
                    "* line {}: found {} cols".format(i, len(row)))
            if not all(
                    v is None or isinstance(v, (str, float, bool, int)) for v in
                    row):
                illegal_value_errs.append("* line {}: {}".format(i, repr(row)))
        errs = []
        if square_errs:
            errs.append(
                "DataArray is not a square (expected {} cols):".format(
                    col_count))
            errs.extend(square_errs)
        if illegal_value_errs:
            errs.append(
                "Found illegal values "
                "(only float, str, None, int, bool are allowed):")
            errs.extend(illegal_value_errs)
        if errs:
            raise ValueError("\n".join(errs))

    def _copy_whole_data_array(
            self, oCell: UnoCell, data_array: DATA_ARRAY):
        row_count = len(data_array)
        col_count = len(data_array[0])
        cell_address = oCell.CellAddress
        column = cell_address.Column
        row = cell_address.Row
        oRange = oCell.Spreadsheet.getCellRangeByPosition(
            column, row, column + col_count - 1, row + row_count - 1)
        oRange.DataArray = data_array
        self._callback(row_count)

    def _copy_data_array_by_chunks(
            self, oCell: UnoCell, data_array: DATA_ARRAY):
        row_count = len(data_array)
        col_count = len(data_array[0])
        cell_address = oCell.CellAddress
        column = cell_address.Column
        row = cell_address.Row

        start_r = 0
        end_r = self._chunk_size
        end_c = column + col_count - 1
        oSheet = oCell.Spreadsheet
        while end_r < row_count:
            oRange = oSheet.getCellRangeByPosition(
                column, row + start_r, end_c, row + end_r - 1)
            oRange.DataArray = data_array[start_r:end_r]
            self._callback(end_r)
            start_r = end_r
            end_r += self._chunk_size

        oRange = oSheet.getCellRangeByPosition(
            column, row + start_r, end_c, row + row_count - 1)
        oRange.DataArray = data_array[start_r:]
        self._callback(row_count)


@contextmanager
def undo_context(oDoc: UnoSpreadsheetDocument, title: Optional[str] = None
                 ) -> ContextManager[None]:
    """
    Do something inside an undo context
    @param oDoc: the document
    @param title: then undo title
    """
    oUndoManager = oDoc.UndoManager
    if title is None:
        oUndoManager.enterHiddenUndoContext()
        try:
            yield
        finally:
            oUndoManager.leaveUndoContext()
    else:
        oUndoManager.enterUndoContext(title)
        try:
            yield
        finally:
            oUndoManager.leaveUndoContext()


@contextmanager
def no_undo_context(oDoc: UnoSpreadsheetDocument) -> ContextManager[None]:
    """
    Do something out of the undo context
    @param oDoc: the document
    """
    oUndoManager = oDoc.UndoManager
    oUndoManager.lock()
    try:
        yield
    finally:
        oUndoManager.unlock()
