#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2025 J. Férard <https://github.com/jferard>
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
import typing
from pathlib import Path
from typing import (Union, Tuple, List, cast, Optional, Sequence, Any)


# Misc
T = typing.TypeVar("T")

# noinspection PyTypeHints
def lazy[T](typ: typing.Type[T]) -> Optional[T]:
    return cast(Optional[typ], None)


#####
# DATA_ARRAY
#####
DATA_VALUE = Union[str, float]
DATA_ROW = Sequence[DATA_VALUE]
DATA_ARRAY = Sequence[DATA_ROW]
StrPath = Union[str, Path]


# BASE
class UnoObject:
    def supportsService(self, name: str) -> bool: ...
    def __repr__(self) -> str: ...

class UnoStruct(UnoObject): ...

class UnoService(UnoObject): ...

class UnoEnum(UnoObject, typing.Generic[T]):
    value: T

######
# structs
######
class UnoRangeAddress(UnoStruct):
    StartColumn: int
    EndColumn: int
    StartRow: int
    EndRow: int
    Sheet: int

class UnoCellAddress(UnoStruct):
    Column: int
    Row: int

class UnoPropertyValue(UnoStruct):
    Name: str

UnoPropertyValues = Union[List[UnoPropertyValue], Tuple[UnoPropertyValue, ...]]

class UnoDateStruct(UnoStruct):
    Year: int
    Month: int
    Day: int

class UnoSizeStruct(UnoStruct):
    Width: int
    Height: int

# Collections
S = typing.TypeVar("S", bound=UnoObject)

class UnoNameAccess(UnoService, typing.Generic[S]):
    ElementNames: Sequence[str]
    def getByName(self, name: str) -> S: ...
    def hasByName(self, name: str) -> bool: ...
    def removeByName(self, name: str) -> None: ...
    def insertNewByName(self, name: str, pos: int, v: Optional[S] = None) -> S: ...

class UnoIndexAccess(UnoService, typing.Generic[S]):
    Count: int
    def getByIndex(self, i: int) -> S: ...
    def removeByIndex(self, i: int, count: int) -> None: ...


class UnoEnumeration(UnoService, typing.Generic[S]):
    def hasMoreElements(self) -> bool: ...
    def nextElement(self) -> S: ...

class UnoEnumerable(UnoService, typing.Generic[S]):
    def createEnumeration(self) -> UnoEnumeration[S]: ...


######
# services
######
class UnoXScriptContext(UnoService):
    def getDocument(self) -> "UnoOfficeDocument": ...
    def getComponentContext(self) -> "UnoContext": ...
    def getDesktop(self) -> "UnoDesktop": ...


class UnoRange(UnoService):
    Spreadsheet: "UnoSheet"
    RangeAddress: "UnoRangeAddress"
    DataArray: DATA_ARRAY
    Size: UnoSizeStruct
    NumberFormat: int
    Rows: UnoIndexAccess["UnoRow"]
    Columns: UnoIndexAccess["UnoColumn"]
    def getCellByPosition(self, c: int, r: int) -> "UnoCell": ...
    def getCellRangeByPosition(self, c1: int, r1: int, c2: int, r2: int) -> "UnoRange": ...
    def createSortDescriptor(self) -> Sequence[UnoPropertyValue]: ...
    def createFilterDescriptor(self, empty: bool) -> Sequence[UnoPropertyValue]: ...
    def sort(self, sort_descriptor: Sequence[UnoPropertyValue]) -> None: ...
    def filter(self, descriptor: Sequence[UnoPropertyValue]) -> None: ...
    def merge(self, m: bool) -> None: ...


class UnoRanges(UnoEnumerable[UnoRange]): ...


class UnoOfficeDocument(UnoService):
    CurrentController: "UnoController"
    StyleFamilies: UnoNameAccess
    URL: str
    CurrentSelection: Union[UnoRange, UnoRanges]
    BasicLibraries: UnoNameAccess

    def getScriptProvider(self) -> UnoObject: ...
    def createInstance(self, name: str) -> UnoObject: ...
    def lockControllers(self) -> None: ...
    def unlockControllers(self) -> None: ...
    def storeAsURL(self, url: str, args: Sequence[UnoPropertyValue]) -> None: ...
    def storeToURL(self, url: str, args: Sequence[UnoPropertyValue]) -> None: ...
    def store(self) -> None: ...
    def close(self, free: bool) -> None: ...


# Calc
class UnoNumberFormats:
    def queryKey(self, fmt: str, oLocale: UnoStruct, create: bool) -> int: ...
    def addNew(self, fmt: str, oLocale: UnoStruct) -> int: ...
    def getByKey(self, format_id: int) -> UnoService: ...
    def getStandardFormat(self, format_id: int, oLocale: UnoStruct) -> int: ...


class UnoUndoManager(UnoService):
    def enterHiddenUndoContext(self) -> None: ...
    def leaveUndoContext(self) -> None: ...
    def enterUndoContext(self, title: str) -> None: ...
    def lock(self) -> None: ...
    def unlock(self) -> None: ...


class UnoSpreadsheetDocument(UnoOfficeDocument):
    Sheets: "UnoSheets"
    NamedRanges: UnoNameAccess
    NumberFormats: UnoNumberFormats
    UndoManager: UnoUndoManager
    NullDate: UnoDateStruct



# deprecated: use UnoSpreadsheetDocument
UnoSpreadsheet = UnoSpreadsheetDocument


class UnoCalcCursor(UnoRange):
    def gotoStartOfUsedArea(self, exp: bool) -> None: ...
    def gotoEndOfUsedArea(self, exp: bool) -> None: ...
    def collapseToMergedArea(self) -> None: ...


class UnoPilotTables(UnoIndexAccess[UnoService], UnoNameAccess[UnoService]):
    def createDataPilotDescriptor(self) -> UnoService: ...


class UnoSheet(UnoRange):
    AbsoluteName: str
    Name: str
    DrawPage: UnoObject
    PageStyle: str
    DataPilotTables: UnoPilotTables
    PivotCharts: UnoNameAccess[UnoService]

    def createCursorByRange(self, oCell: "UnoCell") -> UnoCalcCursor: ...
    def createCursor(self) -> UnoCalcCursor: ...
    def setPrintAreas(self, areas: Sequence[UnoRangeAddress]) -> None: ...
    def setPrintTitleRows(self, bool: bool) -> None: ...
    def setTitleRows(self, title: UnoRangeAddress) -> None: ...
    def copyRange(self, cell_address: UnoCellAddress, range_address: UnoRangeAddress) -> None: ...


class UnoSheets(UnoNameAccess[UnoSheet], UnoIndexAccess[UnoSheet]):
    def copyByName(self, name: str, new_name: str, new_index: int) -> None: ...
    def importSheet(self, doc: UnoSpreadsheetDocument, name: str, dest_position: int): ...


class UnoTextRange(UnoService):
    TextField: UnoService
    String: str
    CharFontName: str
    CharHeight: float
    CharWeight: float
    CharPosture: int
    CharColor: int
    CharBackColor: int
    CharOverline: int
    CharStrikeout: int
    CharUnderline: int
    CharEscapementHeight: float
    CharEscapement: float

    TextPortionType: str

    Start: "UnoTextRange"
    End: "UnoTextRange"

class UnoTextContent(UnoTextRange, UnoEnumerable[UnoTextRange]): ...

class UnoText(UnoTextRange, UnoEnumerable[UnoTextContent]):
    def createTextCursorByRange(self, r: UnoTextRange) -> UnoTextRange: ...


class UnoCell(UnoRange, UnoText):
    String: str
    Value: float
    Formula: str
    Type: UnoEnum[str]
    FormulaResultType: UnoEnum[str]
    CellAddress: UnoCellAddress
    Text: UnoText

    def insertTextContent(self, r: UnoTextRange, content: UnoService, absorb: bool) -> None: ...


class UnoRow(UnoRange): ...

class UnoColumn(UnoRange): ...


# Writer
class UnoTextDocument(UnoOfficeDocument): ...

# Other
class UnoDrawingDocument(UnoOfficeDocument): ...
class UnoPresentationDocument(UnoOfficeDocument): ...


class UnoFrame(UnoService):
    ContainerWindow: UnoObject


class UnoController(UnoService):
    Frame: UnoFrame
    ActiveSheet: UnoSheet

    def select(self, oRanges: Union[UnoRange, UnoRanges]) -> None: ...
    def getTransferable(self) -> UnoService: ...
    def insertTransferable(self, t: UnoService): ...


class UnoContext(UnoService):
    def getServiceManager(self) -> UnoObject: ...
    def getByName(self, _name: str) -> UnoObject: ...

class UnoDesktop(UnoService):
    def getCurrentComponent(self) -> UnoObject: ...
    def loadComponentFromURL(self, url: str, target: str, frame_flags: int, args: Sequence[UnoPropertyValue]) -> UnoOfficeDocument: ...


class UnoDispatcher(UnoService):
    def executeDispatch(self, frame: Union[UnoController, UnoFrame], url: str, target: str, flagt: int, args: Sequence[UnoPropertyValue]) -> None: ...

# Dialogs
class UnoControlModel(UnoService):
    Name: str


class UnoMainControlModel(UnoControlModel):
    def createInstance(self, name: str) -> UnoControlModel: ...
    def insertByName(self, name: str, model: UnoControlModel) -> None: ...
    def getByName(self, name: str) -> UnoControlModel: ...


class UnoControl(UnoService):
    MinimumSize: "UnoSizeStruct"
    Model: UnoControlModel
    def setModel(self, model: UnoControlModel) -> None: ...
    def setFocus(self) -> None: ...
    def setVisible(self, b: bool) -> None: ...
    def addActionListener(self, listener: Any) -> None: ...
    def addKeyListener(self, listener: Any) -> None: ...
    def addItemListener(self, listener: Any) -> None: ...
    def addTextListener(self, listener: Any) -> None: ...


class UnoListControl(UnoControl):
    ItemCount: int
    SelectedItemsPos: List[int]

    def removeItems(self, i: int, c: int) -> None: ...
    def addItems(self, items: List[str], pos: int) -> None: ...
    def selectItemsPos(self, positions: List[int], sel: bool) -> None: ...


class UnoToolkit(UnoService):
    def createMessageBox(
        self, parent_win: UnoControl, msg_type: int, msg_buttons: int, msg_title: str, msg_text: str) -> "UnoMainControl": ...


class UnoMainControl(UnoControl):
    Controls: Sequence[UnoControl]
    def setModel(self, model: UnoMainControlModel) -> None: ...
    def createPeer(self, oToolkit: UnoToolkit, parent_win: Optional["UnoMainControl"]): ...
    def getControl(self, name: str) -> UnoControl: ...
    def execute(self) -> int: ...
    def endExecute(self) -> None: ...
    def dispose(self) -> None: ...


class UnoFilePicker(UnoMainControl):
    SelectedFiles: Sequence[str]
    def initialize(self, desc: Sequence[int]): ...
    def appendFilter(self, title: str, filter: str): ...


class UnoFolderPicker(UnoMainControl):
    Title: str
    DisplayDirectory: str
    Directory: str
    def initialize(self, desc: Sequence[int]): ...
    def appendFilter(self, title: str, filter: str): ...

##
# BASE
#

class UnoDatabaseDocument(UnoOfficeDocument): 
    ...



class UnoDBStatement(UnoService):
    def executeUpdate(self, sql: str) -> None: ...
    def execute(self, sql: str) -> None: ...
    def addBatch(self, sql: str) -> None: ...
    def executeBatch(self) -> None: ...


class UnoDBDrop(UnoNameAccess[T], UnoIndexAccess[T], typing.Generic[T]):
    def dropByName(self, name: str) -> None: ...
    def dropByIndex(self, i: int) -> None: ...


class UnoDBTable(UnoService):
    Keys: UnoDBDrop[UnoService]
    Indexes: UnoDBDrop[UnoService]

class UnoDBTables(UnoDBDrop[UnoDBTable]):
    def createDataDescriptor(self) -> UnoService: ...
    def appendByDescriptor(self, desc: UnoService) -> None: ...


class UnoDBConnection(UnoService):
    Tables: UnoDBTables
    Views: UnoDBDrop[UnoService]

    def isClosed(self) -> bool: ...
    def createStatement(self) -> UnoDBStatement: ...
    def commit(self) -> None: ...
    def close(self) -> None: ...


class UnoDBAccess(UnoService):
    DatabaseDocument: UnoDatabaseDocument
    QueryDefinitions: Union[UnoIndexAccess, UnoNameAccess]

    def connectWithCompletion(self, oHandler: UnoService) -> UnoDBConnection: ...


class UnoDBContext(UnoNameAccess[UnoDBAccess]):
    def createInstance(self) -> UnoDBAccess: ...


## CB
class UnoDataFlavor(UnoStruct):
    MimeType: Tuple[str, str]

class UnoTransferable(UnoService):
    def getTransferDataFlavors(self) -> List[UnoDataFlavor]: ...
    def getTransferData(self, t: UnoStruct) -> Any: ...


class UnoClipboard(UnoService):
    def setContents(self, param, param1): ...
    def getContents(self) -> UnoTransferable: ...