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
import enum
import os
from pathlib import Path
from urllib.parse import urlparse


##############################
# # DO NOT EMBED THIS FILE # #
##############################
class DataType(enum.Enum):
    BIT = -7
    TINYINT = -6
    SMALLINT = 5
    INTEGER = 4
    BIGINT = -5
    FLOAT = 6
    REAL = 7
    DOUBLE = 8
    NUMERIC = 2
    DECIMAL = 3
    CHAR = 1
    VARCHAR = 12
    LONGVARCHAR = -1
    DATE = 91
    TIME = 92
    TIMESTAMP = 93
    BINARY = -2
    VARBINARY = -3
    LONGVARBINARY = -4
    SQLNULL = 0
    OTHER = 1111
    OBJECT = 2000
    DISTINCT = 2001
    STRUCT = 2002
    ARRAY = 2003
    BLOB = 2004
    CLOB = 2005
    REF = 2006
    BOOLEAN = 16


class ColumnValue:
    pass


class MessageBoxType:
    MESSAGEBOX = 0
    INFOBOX = 1
    WARNINGBOX = 2
    ERRORBOX = 3
    QUERYBOX = 4


class MessageBoxButtons:
    BUTTONS_OK = 1
    BUTTONS_OK_CANCEL = 2
    BUTTONS_YES_NO = 3
    BUTTONS_YES_NO_CANCEL = 4
    BUTTONS_RETRY_CANCEL = 5
    BUTTONS_ABORT_IGNORE_RETRY = 6
    DEFAULT_BUTTON_OK = 0x10000
    DEFAULT_BUTTON_CANCEL = 0x20000
    DEFAULT_BUTTON_RETRY = 0x30000
    DEFAULT_BUTTON_YES = 0x40000
    DEFAULT_BUTTON_NO = 0x50000
    DEFAULT_BUTTON_IGNORE = 0x60000


class MessageBoxResults:
    CANCEL = 0
    OK = 1
    YES = 2
    NO = 3
    RETRY = 4
    IGNORE = 5


class ExecutableDialogResults:
    OK = 1,
    CANCEL = 0


class FontWeight:
    BOLD = None


class PushButtonType:
    OK = None,
    CANCEL = None


class XTransferable:
    pass


class FrameSearchFlag:
    AUTO = None


class BorderLineStyle:
    SOLID = 0


class ConditionOperator:
    FORMULA = 9


class ValidationType:
    LIST = 6


class TableValidationVisibility:
    SORTEDASCENDING = None
    UNSORTED = None


class ScriptFrameworkErrorException(Exception):
    pass


class UnoRuntimeException(Exception):
    pass


class UnoException(Exception):
    pass


Locale = object


class NumberFormat:
    ALL = 0
    DEFINED = 1
    DATE = 2
    TIME = 4
    CURRENCY = 8
    NUMBER = 16
    SCIENTIFIC = 32
    FRACTION = 64
    PERCENT = 128
    TEXT = 256
    DATETIME = 6
    LOGICAL = 1024
    UNDEFINED = 2048
    EMPTY = 4096
    DURATION = 8196


class uno:
    @staticmethod
    def fileUrlToSystemPath(url: str) -> str:
        result = urlparse(url)
        if result.netloc:
            return os.path.join(result.netloc, result.path.lstrip("/"))
        else:
            return result.path

    @staticmethod
    def systemPathToFileUrl(path: str) -> str:
        return Path(path).as_uri()

    @staticmethod
    def createUnoStruct(struct_id: str) -> None:
        pass

    @staticmethod
    def Any(name, value) -> None:
        pass


class unohelper:
    class Base:
        pass

    @staticmethod
    def ImplementationHelper():
        class C:
            @staticmethod
            def addImplementation(*args): return None

        return C


class PropertyState:
    DIRECT_VALUE = 0
    DEFAULT_VALUE = 1
    AMBIGUOUS_VALUE = 2


class FontSlant:
    NONE = 0
    OBLIQUE = 1
    ITALIC = 2
