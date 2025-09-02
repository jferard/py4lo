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
import enum


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


class FontWeight:
    BOLD = 150


class ExecutableDialogResults:
    CANCEL = 0
    OK = 1


class PushButtonType:
    OK = None
    CANCEL = None


class FrameSearchFlag:
    AUTO = 0


class BorderLineStyle:
    SOLID = 0


class ConditionOperator:
    FORMULA = 9


class ValidationType:
    LIST = 6


class TableValidationVisibility:
    SORTEDASCENDING = 2
    UNSORTED = 1


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


class PropertyState:
    DIRECT_VALUE = 0
    DEFAULT_VALUE = 1
    AMBIGUOUS_VALUE = 2


class FontSlant:
    NONE = 0
    OBLIQUE = 1
    ITALIC = 2


class WindowClass:
    TOP = 0
    MODALTOP = 1
    CONTAINER = 2
    SIMPLE = 3


class WindowAttribute:
    SHOW = 1
    FULLSIZE = 2
    OPTIMUMSIZE = 4
    MINSIZE = 8
    BORDER = 16
    SIZEABLE = 32
    MOVEABLE = 64
    CLOSEABLE = 128
    SYSTEMDEPENDENT = 256
    NODECORATION = 512


class PosSize:
    X = 1
    Y = 2
    WIDTH = 4
    HEIGHT = 8
    POS = 3
    SIZE = 12
    POSSIZE = 15


class ScrollBarOrientation:
    HORIZONTAL = 0
    VERTICAL = 1


class GeneralFunction2:
    NONE = 0
    AUTO = 1
    SUM = 2
    COUNT = 3
    AVERAGE = 4
    MAX = 5
    MIN = 6
    PRODUCT = 7
    COUNTNUMS = 8
    STDEV = 9
    STDEVP = 10
    VAR = 11
    VARP = 12
    MEDIAN = 13


class DataPilotOutputRangeType:
    WHOLE = 0
    TABLE = 1
    RESULT = 2


class CellContentType:
    EMPTY = 0
    VALUE = 1
    TEXT = 2
    FORMULA = 3


class TypeClass(object):
    VOID = 0
    CHAR = 1
    BOOLEAN = 2
    BYTE = 3
    SHORT = 4
    UNSIGNED_SHORT = 5
    LONG = 6
    UNSIGNED_LONG = 7
    HYPER = 8
    UNSIGNED_HYPER = 9
    FLOAT = 10
    DOUBLE = 11
    STRING = 12
    TYPE = 13
    ANY = 14
    ENUM = 15
    TYPEDEF = 16
    STRUCT = 17
    EXCEPTION = 19
    SEQUENCE = 20
    INTERFACE = 22
    SERVICE = 23
    MODULE = 24
    INTERFACE_METHOD = 25
    INTERFACE_ATTRIBUTE = 26
    UNKNOWN = 27
    PROPERTY = 28
    CONSTANT = 29
    CONSTANTS = 30
    SINGLETON = 31


class TemplateDescription:
    """https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1ui_1_1dialogs_1_1TemplateDescription.html"""
    FILEOPEN_SIMPLE = 0
    FILESAVE_SIMPLE = 1
    FILESAVE_AUTOEXTENSION_PASSWORD = 2
    FILESAVE_AUTOEXTENSION_PASSWORD_FILTEROPTIONS = 3
    FILESAVE_AUTOEXTENSION_SELECTION = 4
    FILESAVE_AUTOEXTENSION_TEMPLATE = 5
    FILEOPEN_LINK_PREVIEW_IMAGE_TEMPLATE = 6
    FILEOPEN_PLAY = 7
    FILEOPEN_READONLY_VERSION = 8
    FILEOPEN_LINK_PREVIEW = 9
    FILESAVE_AUTOEXTENSION = 10
    FILEOPEN_PREVIEW = 11
    FILEOPEN_LINK_PLAY = 12
    FILEOPEN_LINK_PREVIEW_IMAGE_ANCHOR = 13
