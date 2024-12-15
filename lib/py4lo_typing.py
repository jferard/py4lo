#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2024 J. Férard <https://github.com/jferard>
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
from typing import (NewType, Any, Union, Tuple, List, cast, Optional)

UnoXScriptContext = NewType("UnoXScriptContext", Any)

UnoObject = NewType("UnoObject", Any)
UnoStruct = NewType("UnoStruct", UnoObject)
UnoEnum = NewType("UnoEnum", UnoObject)
UnoService = NewType("UnoService", UnoObject)

######
# services
######
UnoOfficeDocument = NewType("UnoOfficeDocument", UnoService)

# Calc
UnoSpreadsheetDocument = NewType("UnoSpreadsheetDocument", UnoOfficeDocument)
# deprecated: use UnoSpreadsheetDocument
UnoSpreadsheet = UnoSpreadsheetDocument
UnoRange = NewType("UnoRange", UnoService)
UnoSheet = NewType("UnoSheet", UnoRange)
UnoCell = NewType("UnoCell", UnoRange)
UnoRow = NewType("UnoRow", UnoRange)
UnoColumn = NewType("UnoColumn", UnoRange)
UnoCellAddress = NewType("UnoCellAddress", UnoService)
UnoTextRange = NewType("UnoTextRange", UnoService)

# Writer
UnoTextDocument = NewType("UnoTextDocument", UnoOfficeDocument)

# Other
UnoDrawingDocument = NewType("UnoDrawingDocument", UnoOfficeDocument)
UnoPresentationDocument = NewType("UnoPresentationDocument", UnoOfficeDocument)

UnoController = NewType("UnoController", UnoService)
UnoContext = NewType("UnoContext", UnoService)

UnoControlModel = NewType("UnoControlModel", UnoService)
UnoControl = NewType("UnoControl", UnoService)

######
# structs
######
UnoRangeAddress = NewType("UnoRangeAddress", UnoStruct)
UnoPropertyValue = NewType("UnoPropertyValue", UnoStruct)
UnoPropertyValues = Union[List[UnoPropertyValue], Tuple[UnoPropertyValue, ...]]

#####
# DATA_ARRAY
#####
DATA_VALUE = Union[str, float]
DATA_ROW = Union[Tuple[DATA_VALUE, ...], List[DATA_VALUE]]
DATA_ARRAY = Union[Tuple[DATA_ROW, ...], List[DATA_ROW]]
StrPath = Union[str, Path]

# Misc
def lazy(typ):
    return cast(Optional[typ], None)