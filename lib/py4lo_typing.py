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
def lazy(typ) -> Optional[Any]:
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
    def insertNewByName(self, name: str, pos) -> S: ...

class UnoIndexAccess(UnoService, typing.Generic[S]):
    Count: int
    def getByIndex(self, i: int) -> S: ...
    def removeByIndex(self, i: int, count: int) -> None: ...


class UnoEnum(UnoService, typing.Generic[S]):
    def hasMoreElements(self) -> bool: ...
    def nextElement(self) -> S: ...

class UnoEnumerable(UnoService, typing.Generic[S]):
    def createEnumeration(self) -> UnoEnum[S]: ...


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


class UnoRanges(UnoService):
    pass


class UnoOfficeDocument(UnoService):
    CurrentController: "UnoController"
    StyleFamilies: UnoNameAccess
    URL: str
    CurrentSelection: Union[UnoRange, UnoRanges]
    def getScriptProvider(self) -> UnoObject: ...
    def createInstance(self, name: str) -> UnoObject: ...
    def lockControllers(self) -> None: ...
    def unlockControllers(self) -> None: ...
    def storeAsURL(self, url: str, args: Sequence[UnoPropertyValue]) -> None: ...
    def storeToURL(self, url: str, args: Sequence[UnoPropertyValue]) -> None: ...
    def close(self, free: bool) -> None: ...


# Calc
class UnoNumberFormats:
    def queryKey(self, fmt: str, oLocale: UnoStruct, create: bool) -> int: ...
    def addNew(self, fmt: str, oLocale: UnoStruct) -> int: ...
    def getByKey(self, format_id: int) -> UnoService: ...


class UnoSpreadsheetDocument(UnoOfficeDocument):
    Sheets: "UnoSheets"
    NamedRanges: UnoNameAccess
    NumberFormats: UnoNumberFormats



# deprecated: use UnoSpreadsheetDocument
UnoSpreadsheet = UnoSpreadsheetDocument


class UnoCalcCursor(UnoRange):
    def gotoStartOfUsedArea(self, exp: bool) -> None: ...
    def gotoEndOfUsedArea(self, exp: bool) -> None: ...
    def collapseToMergedArea(self) -> None: ...


class UnoSheet(UnoRange):
    AbsoluteName: str
    Name: str
    DrawPage: UnoObject
    PageStyle: str
    DataPilotTables: UnoIndexAccess[UnoService]

    def createCursorByRange(self, oCell: "UnoCell") -> UnoCalcCursor: ...
    def createCursor(self) -> UnoCalcCursor: ...
    def setPrintAreas(self, areas: Sequence[UnoRangeAddress]) -> None: ...
    def setPrintTitleRows(self, bool: bool) -> None: ...
    def setTitleRows(self, title: UnoRangeAddress) -> None: ...
    def copyRange(self, cell_address: UnoCellAddress, range_address: UnoRangeAddress) -> None: ...


class UnoSheets(UnoNameAccess[UnoSheet], UnoIndexAccess[UnoSheet]):
    def copyByName(self, name: str, new_name: str, new_index: int) -> None: ...


class UnoTextRange(UnoService): ...

class UnoText(UnoEnumerable[UnoTextRange]): ...

class UnoCell(UnoRange, UnoText):
    String: str
    Type: UnoObject
    FormulaResultType: UnoObject
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
class UnoControlModel(UnoService): ...

class UnoMainControlModel(UnoControlModel):
    def createInstance(self, name: str) -> UnoControlModel: ...
    def insertByName(self, name: str, model: UnoControlModel) -> None: ...

class UnoControl(UnoService):
    MinimumSize: "UnoSizeStruct"
    Model: UnoControlModel
    def setModel(self, model: UnoControlModel) -> None: ...
    def setFocus(self) -> None: ...
    def setVisible(self, b: bool) -> None: ...


class UnoToolkit(UnoService):
    def createMessageBox(
        self, parent_win: UnoControl, msg_type: int, msg_buttons: int, msg_title: str, msg_text: str) -> "UnoMainControl": ...


class UnoMainControl(UnoControl):
    def setModel(self, model: UnoMainControlModel) -> None: ...
    def execute(self) -> int: ...
    def createPeer(self, oToolkit: UnoToolkit, parent_win: Optional["UnoMainControl"]): ...
    def getControl(self, name: str) -> UnoControl: ...
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
