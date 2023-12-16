#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
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
"""
Basic support for uno types out of the LibreOffice engine.
"""
# mypy: disable-error-code="empty-body"
from pathlib import Path
from typing import (NewType, Any, Union, Tuple, List, TypeVar, Generic, cast,
                    Optional, Collection)

#####
# DATA_ARRAY
#####
DATA_VALUE = Union[str, float]
DATA_ROW = Union[Tuple[DATA_VALUE, ...], List[DATA_VALUE]]
DATA_ARRAY = Union[Tuple[DATA_ROW, ...], List[DATA_ROW]]

# str path
StrPath = Union[str, Path]


#####
# base types
#####
class UnoObject: ...


class UnoStruct(UnoObject): ...


class UnoEnum(UnoObject):
    value: Any


class UnoService(UnoObject):
    def supportsService(self, _: str) -> bool: ...



# type vars
UO = TypeVar("UO", bound=UnoObject)


######
# structs
######
class UnoCellRangeAddress(UnoStruct):
    Sheet: int
    StartColumn: int
    StartRow: int
    EndColumn: int
    EndRow: int


class UnoPropertyValue(UnoStruct):
    Handle: int
    State: Any  # com::sun::star::beans::PropertyState
    Name: str
    Value: Any


UnoPropertyValues = Union[List[UnoPropertyValue], Tuple[UnoPropertyValue, ...]]


class UnoSize(UnoStruct):
    Height: int
    Width: int


####
# Factories
####
class UnoMultiServiceFactory(UnoService):
    def createInstance(self, _sname: str) -> UnoService: ...


######
# services
######
class UnoOfficeDocument(UnoMultiServiceFactory):
    CurrentController: "UnoController"
    StyleFamilies: "UnoStyleFamilies"
    URL: str

    def getScriptProvider(self) -> "UnoScriptProvider": ...

    def lockControllers(self): ...

    def unlockControllers(self): ...

    def close(self, _b: bool): ...

    def storeAsURL(self, _url: str, _args: UnoPropertyValues): ...

    def storeToURL(self, _url: str, _args: UnoPropertyValues): ...


# Calc
class UnoConditionalFormat(UnoService):
    def clear(self): ...

    def addNew(self, _: Tuple[UnoPropertyValue, ...]): ...


class UnoFilterDescriptor(UnoService): ...

class UnoCharacterProperties(UnoService):
    CharFontName: str
    CharHeight: float
    CharWeight: float
    CharWeightAsian: float
    CharWeightComplex: float
    CharPosture: int
    CharColor: int
    CharBackColor: int
    CharOverline: int
    CharStrikeout: int
    CharUnderline: int
    CharEscapementHeight: float
    CharEscapement: float


class UnoRange(UnoCharacterProperties):
    Spreadsheet: "UnoSheet"
    Columns: "UnoTableColumns"
    Rows: "UnoTableRows"
    DataArray: DATA_ARRAY
    Size: UnoSize

    IsTextWrapped: bool
    RangeAddress: "UnoCellRangeAddress"
    ConditionalFormat: UnoConditionalFormat
    NumberFormat: int

    def getCellByPosition(self, _c: int, _r: int) -> "UnoCell": ...

    def getCellRangeByPosition(self, _c1: int, _r1: int, _c2: int,
                               _r2: int) -> "UnoRange": ...

    def createSortDescriptor(self) -> List[UnoPropertyValue]: ...

    def sort(self, _desc: List[UnoPropertyValue]): ...

    def createFilterDescriptor(self, _b: bool): UnoFilterDescriptor

    def filter(self, _: UnoFilterDescriptor): ...


class UnoValidationPS(UnoService):
    Type: int
    IgnoreBlankCells: bool
    ShowList: int
    ShowErrorMessage: bool
    Formula1: str

class UnoTextRange(UnoService): ...
class UnoTextCursor(UnoService): ...

class UnoTextURL(UnoService):
    Representation: str
    URL: str

class UnoCell(UnoRange):
    Type: UnoEnum
    FormulaResultType: UnoEnum
    String: str
    Value: float

    Validation: UnoValidationPS
    Text: "UnoText"

    def insertTextContent(self, _: UnoTextCursor, _a, _b: bool): ... # TODO

class UnoRow(UnoRange):
    Height: int
    OptimalHeight: int


class UnoColumn(UnoRange):
    Width: int
    OptimalWidth: int


class UnoDatabaseDocument(UnoOfficeDocument): ...



class UnoSpreadsheetDocument(UnoOfficeDocument):
    NumberFormats: "UnoNumberFormats"
    NamedRanges: "UnoNamedRanges"  # TODO
    Sheets: "UnoSheets"  # TODO



# deprecated: use UnoSpreadsheetDocument
UnoSpreadsheet = UnoSpreadsheetDocument


class UnoNumberFormat(UnoService):
    Type: int


class UnoNumberFormats(UnoService):
    def getByKey(self, _key: int) -> UnoNumberFormat: ...
    def getStandardFormat(self, _id: Any, _locale: Any): ...
    def queryKey(self, _fmt_name: str, _locale: UnoStruct,
                 _param: bool) -> int: ...

    def addNew(self, _fmt_name: str, _locale: UnoStruct) -> int: ...


class UnoXScriptContext:
    def getDocument(self) -> "UnoOfficeDocument": ...

    def getComponentContext(self) -> "UnoComponentContext": ...

    def getDesktop(self) -> "UnoDesktop": ...


class UnoNameAccess(Generic[UO], UnoService):
    ElementNames: List[str]

    def getByName(self, _name: str) -> UO: ...
    def hasByName(self, _name: str) -> bool: ...

    def removeByName(self, _name: str): ...

    def insertByName(self, _name: str, _o: UO): ...


class UnoNamedRange(UnoService):
    ReferredCells: UnoRange


class UnoNamedRanges(UnoNameAccess[UnoNamedRange]): ...


class UnoIndexAccess(Generic[UO], UnoService):
    Count: int

    def getByIndex(self, _i: int) -> UO: ...

    def removeByIndex(self, _i: int): ...


class UnoEnumeration(Generic[UO], UnoService):
    def hasMoreElements(self) -> bool: ...

    def nextElement(self) -> UO: ...


class UnoEnumerationAccess(Generic[UO], UnoService):
    def createEnumeration(self) -> UnoEnumeration[UO]: ...


class UnoSheetCellCursor(UnoRange):
    def collapseToMergedArea(self): ...

    def gotoStartOfUsedArea(self, _param: bool): ...

    def gotoEndOfUsedArea(self, _param: bool): ...


class UnoSheet(UnoRange, UnoNameAccess, UnoIndexAccess, UnoEnumerationAccess):
    DrawPage: "UnoDrawPage"
    PageStyle: str
    Name: str

    def createCursorByRange(self, _oRange: UnoRange) -> UnoSheetCellCursor: ...

    def createCursor(self) -> UnoSheetCellCursor: ...

    def setPrintAreas(self, _: Collection[UnoCellRangeAddress]): ...

    def setPrintTitleRows(self, _: bool): ...

    def setTitleRows(self, _: UnoCellRangeAddress): ...

class UnoSheets(UnoIndexAccess[UnoSheet], UnoNameAccess[UnoSheet]):
    def insertNewByName(self, _name: str, _s: int): ...

    def copyByName(self, _name1: str, _name2: str, _s: int): ...

    def importSheet(self, _source: UnoSpreadsheetDocument, _name: str, _dest_position: int): ...


class UnoTableColumns(UnoEnumerationAccess[UnoColumn],
                      UnoIndexAccess[UnoColumn]): ...


class UnoTableRows(UnoEnumerationAccess[UnoRow], UnoIndexAccess[UnoRow]): ...


class UnoDrawPage(UnoService):
    Forms: "UnoForms"


class UnoMultiComponentFactory(UnoService):
    def createInstanceWithContext(self, _serviceSpecifier: str,
                                  _context: "UnoComponentContext") -> UnoService: ...

    def createInstanceWithArgumentsAndContext(
            self, _serviceSpecifier: str, _args: List[Any],
            _context: "UnoComponentContext") -> UnoService: ...

    def getAvailableServiceNames(self) -> List[str]: ...


class UnoComponentContext(UnoService):
    def getServiceManager(self) -> "UnoServiceManager":
        pass


class UnoForms(UnoService):
    Parent: "UnoSpreadsheetDocument"


class UnoCellAddress(UnoService):
    Column: int
    Row: int


class UnoTextField(UnoService):
    URL: str
    Representation: str


class UnoText(UnoEnumerationAccess["UnoText"], UnoCharacterProperties):
    TextPortionType: int
    TextField: UnoTextField
    String: str

    End: UnoTextRange

    def createTextCursorByRange(self, _: UnoTextRange) -> UnoTextCursor: ...


# Writer
UnoTextDocument = NewType("UnoTextDocument", UnoOfficeDocument)

# Other
UnoDrawingDocument = NewType("UnoDrawingDocument", UnoOfficeDocument)
UnoPresentationDocument = NewType("UnoPresentationDocument", UnoOfficeDocument)


class UnoController(UnoService):
    Frame: "UnoFrame"
    ActiveSheet: "UnoSheet"

    def select(self, _: Any): ...


class UnoFrame(UnoService):
    ContainerWindow: "UnoWindow"


#####
# Styles
#####
class UnoStyle(UnoService): ...


class UnoPageStyle(UnoStyle):
    IsLandscape: bool
    ScaleToPagesX: int
    ScaleToPagesY: int
    Size: UnoSize


class UnoStyleFamily(UnoIndexAccess[UnoStyle], UnoNameAccess[UnoStyle]): ...


class UnoStyleFamilies(UnoIndexAccess[UnoStyleFamily],
                       UnoNameAccess[UnoStyleFamily]): ...


####
# Scripts
####

class DataFlavor(UnoService):
    MimeType: Tuple[str, str]
class Transferable(UnoService):
    def getTransferDataFlavors(self) -> List[DataFlavor]: ...

    def getTransferData(self, _type: DataFlavor): ...


class UnoClipboard(UnoService):
    def setContents(self, _trans: Transferable, _owner: Any): ...

    def getContents(self) -> Transferable:...


class UnoScript(UnoService):
    ...
class UnoScriptProvider(UnoService):
    def getScript(self, _: str) -> UnoScript: ...


class UnoServiceManager(UnoMultiServiceFactory, UnoMultiComponentFactory): ...


class UnoDesktop(UnoService):
    def loadComponentFromURL(self, _url: str, _target: str, _frame_flags: Any,
                             params: UnoPropertyValues) -> UnoOfficeDocument: ...


class UnoPosSize(UnoStruct):
    Width: int
    Height: int

class UnoWindow(UnoService):
    PosSize: UnoPosSize


UnoContext = NewType("UnoContext", UnoService)

class UnoControlModel(UnoNameAccess["UnoControlModel"]):
    Title: str
    Label: str
    PositionX: int
    PositionY: int
    Width: int
    Height: int
    def createInstance(self, _: str) -> "UnoControlModel": ...

class UnoTextControlModel(UnoControlModel):
    NoLabel: bool

class UnoButtonControlModel(UnoControlModel):
    PushButtonType: Any
    DefaultButton: bool

class UnoFilePicker(UnoService):
    Title: str
    Directory: str
    DisplayDirectory: str
    SelectedFiles: List[str]
    MultiSelectionMode: bool

    def execute(self) -> Any: ...

    def appendFilter(self, _title: str, _filter: str): ...


class UnoFolderPicker(UnoService):
    Title: str
    Directory: str
    DisplayDirectory: str

    def execute(self) -> Any: ...


class UnoToolKit(UnoService):
    def createMessageBox(self, _win: UnoWindow, msg_type: int, msg_buttons: int, msg_title: str,
                         msg_text: str): ...

class UnoControl(UnoService):
    Value: int
    def setModel(self, _m: UnoControlModel): ...

    def createPeer(self, _toolkit: UnoService, _win: Optional[UnoWindow]): ...

    def getControl(self, _n: str) -> "UnoControl": ...

    def execute(self): ...

    def dispose(self): ...

    def setFocus(self): ...

    def setSelection(self, _s: UnoStruct): ...

    def setVisible(self, _b: bool): ...

class UnoTextControl(UnoControl):
    Text: str
    MinimumSize: UnoPosSize

    def insertText(self, _s: UnoStruct, _t: str): ...


### BASE
class UnoStatement(UnoService):
    def executeUpdate(self, _sql: str): ...

    def execute(self, _sql: str): ...

    def addBatch(self, _sql: str): ...

    def executeBatch(self): ...


class UnoContainer(Generic[UO], UnoNameAccess[UO], UnoIndexAccess[UO], UnoEnumerationAccess[UO]):

    def createDataDescriptor(self): ...

    def appendByDescriptor(self, _oTableDescriptor): ...

    def dropByName(self, _name: str): ...
    def dropByIndex(self, _i: int): ...


class UnoConnection(UnoService):
    Views : UnoContainer
    Tables: UnoContainer

    def createStatement(self) -> UnoStatement: ...

    def commit(self): ...

    def close(self): ...


class UnoTable(UnoService):
    Keys: UnoContainer
    Indexes: UnoContainer
class UnoQuery(UnoService):
    Command: str

class UnoQueryDefinitions(UnoNameAccess[UnoQuery]):
    def createInstance(self) -> UnoQuery: ...


class UnoDocumentDataSource(UnoService):
    DatabaseDocument: UnoDatabaseDocument
    QueryDefinitions: UnoQueryDefinitions

    def connectWithCompletion(self, _oHandler: UnoService): ...

class UnoDatabaseContext(UnoContainer):
    def createInstance(self) -> UnoDocumentDataSource:...

def lazy(typ):
    return cast(Optional[typ], None)
