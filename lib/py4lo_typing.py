#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2022 J. Férard <https://github.com/jferard>
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

from typing import (NewType, Any, Union, Tuple, List)

UnoObject = NewType("UnoObject", Any)
UnoStruct = NewType("UnoStruct", Any)

UnoSpreadsheet = NewType("UnoSpreadsheet", UnoObject)
UnoRange = NewType("UnoRange", UnoObject)
UnoRangeAddress = NewType("UnoRangeAddress", UnoStruct)
UnoSheet = NewType("UnoSheet", UnoRange)
UnoCell = NewType("UnoCell", UnoRange)
UnoCellAddress = NewType("UnoCellAddress", UnoObject)


UnoController = NewType("UnoController", UnoObject)
UnoContext = NewType("UnoContext", UnoObject)
UnoService = NewType("UnoService", UnoObject)

UnoPropertyValue = NewType("UnoPropertyValue", UnoStruct)


DATA_VALUE = Union[str, float]
DATA_ROW = Union[Tuple[DATA_VALUE, ...], List[DATA_VALUE]]
DATA_ARRAY = Union[Tuple[DATA_ROW, ...], List[DATA_ROW]]
