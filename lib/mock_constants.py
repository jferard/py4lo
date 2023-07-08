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


############################
## DO NOT EMBED THIS FILE ##
############################
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
    MESSAGEBOX = None


class MessageBoxButtons:
    BUTTONS_OK = None


class FontWeight:
    BOLD = None

class ExecutableDialogResults:
    OK = None,
    CANCEL = None


class PushButtonType:
    OK = None,
    CANCEL = None

class XTransferable:
    pass

class FrameSearchFlag:
    AUTO = None


class BorderLineStyle:
    SOLID = None


class ConditionOperator:
    FORMULA = None


class ValidationType:
    LIST = None


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
    pass

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


class unohelper:
    class Base:
        pass