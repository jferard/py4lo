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
A wrapper for the C SQLite3 API (sqlite module is not shipped with LibreOffice).
The API is **not** compliant to the PEP249 (https://peps.python.org/pep-0249/).

Example:
```
with sqlite_open(self._path, "crw") as db:
    db.execute_update(
        "CREATE TABLE identifiers(identifier INTEGER, name TEXT)")

    with db.transaction():
        with db.prepare("INSERT INTO identifiers VALUES(?, ?)") as stmt:
            stmt.reset()
            stmt.clear_bindings()
            stmt.bind_int(1, 1)
            stmt.bind_text(2, "foo"")
            try:
                self.assertEqual(1, stmt.execute_update())
            except Exception as e:
                print(e)

    with db.prepare("SELECT * FROM identifiers") as stmt:
        for db_row in stmt.execute_query():
            print(db_row)
```

If the library is not in the standard directories, use the environment
variable SQLITE3_LIB to set the actual library path (including the name).
"""

import enum
import os
from contextlib import contextmanager
from ctypes import (
    cdll, c_void_p, byref, c_int, c_char_p, POINTER, c_double, string_at, CDLL
)
from ctypes.util import find_library
from pathlib import Path
from typing import Union, Generator, List, Any, Iterator, Mapping

library_name = find_library('sqlite3')
if library_name is None:
    path = Path.cwd() / "sqlite3.dll" # windows
    if path.exists():
        str_path = str(path)
    else:
        path = next(Path.cwd().glob("libsqlite3.so*")) # linux / mac
        if path is None:
            str_path = os.environ[
                "SQLITE3_LIB"
            ]  # will raise an error if not present
        else:
            str_path = str(path)
    sqlite3_lib = CDLL(str_path)
else:
    sqlite3_lib = cdll.LoadLibrary(library_name)

# opaque structure
sqlite3_p = c_void_p
sqlite3_stmt_p = c_void_p


class SQLiteError(Exception):
    """
    A SQLite error representation: a result code and a message
    """
    def __init__(self, result_code: int, msg: str):
        """
        @param result_code: the result code of the command (see.
        https://www.sqlite.org/rescode.html)
        @param msg: the associated message
        """
        self.result_code = result_code
        self.msg = msg

    def __repr__(self) -> str:
        return "SQLiteError({}, {})".format(self.result_code, repr(self.msg))


class TransactionMode(enum.Enum):
    """
    See https://sqlite.org/lang_transaction.html.
    """
    DEFERRED = "DEFERRED"
    IMMEDIATE = "IMMEDIATE"
    EXCLUSIVE = "EXCLUSIVE"


##############################################
# https://www.sqlite.org/c3ref/funclist.html
##############################################
# https://www.sqlite.org/c3ref/open.html
sqlite3_lib.sqlite3_open_v2.argtypes = [c_char_p, POINTER(sqlite3_p), c_int,
                                        c_char_p]
sqlite3_lib.sqlite3_open_v2.restype = c_int
sqlite3_open_v2 = sqlite3_lib.sqlite3_open_v2

# https://www.sqlite.org/c3ref/errcode.html
sqlite3_lib.sqlite3_errmsg.argtypes = [sqlite3_p]
sqlite3_lib.sqlite3_errmsg.restype = c_char_p
sqlite3_errmsg = sqlite3_lib.sqlite3_errmsg

# https://www.sqlite.org/c3ref/close.html
try:
    sqlite3_close_v2 = sqlite3_lib.sqlite3_close_v2
except AttributeError:
    sqlite3_close_v2 = sqlite3_lib.sqlite3_close

sqlite3_close_v2.argtypes = [sqlite3_p]
sqlite3_close_v2.restype = c_int

# https://www.sqlite.org/c3ref/busy_timeout.html
sqlite3_lib.sqlite3_busy_timeout.argtypes = [sqlite3_p, c_int]
sqlite3_lib.sqlite3_busy_timeout.restype = c_int
sqlite3_busy_timeout = sqlite3_lib.sqlite3_busy_timeout

# https://sqlite.org/c3ref/interrupt.html
sqlite3_lib.sqlite3_interrupt.argtypes = [sqlite3_p]
sqlite3_lib.sqlite3_interrupt.restype = None
sqlite3_interrupt = sqlite3_lib.sqlite3_interrupt

try:
    sqlite3_lib.sqlite3_is_interrupted.argtypes = [sqlite3_p]
    sqlite3_lib.sqlite3_is_interrupted.restype = c_int
    sqlite3_is_interrupted = sqlite3_lib.sqlite3_is_interrupted
except AttributeError:
    def sqlite3_is_interrupted(db):
        return -1

# https://www.sqlite.org/c3ref/exec.html
sqlite3_lib.sqlite3_exec.argtypes = [sqlite3_p, c_char_p, c_void_p, c_void_p,
                                     POINTER(c_char_p)]
sqlite3_lib.sqlite3_exec.restype = c_int
sqlite3_exec = sqlite3_lib.sqlite3_exec

# https://www.sqlite.org/c3ref/prepare.html
sqlite3_lib.sqlite3_prepare_v2.argtypes = [sqlite3_p, c_char_p, c_int,
                                           POINTER(sqlite3_stmt_p),
                                           POINTER(c_char_p)]
sqlite3_lib.sqlite3_prepare_v2.restype = c_int
sqlite3_prepare_v2 = sqlite3_lib.sqlite3_prepare_v2

# https://www.sqlite.org/c3ref/column_count.html
sqlite3_lib.sqlite3_column_count.argtypes = [sqlite3_stmt_p]
sqlite3_lib.sqlite3_column_count.restype = c_int
sqlite3_column_count = sqlite3_lib.sqlite3_column_count

# https://www.sqlite.org/c3ref/bind_blob.html
sqlite3_lib.sqlite3_bind_blob.argtypes = [sqlite3_stmt_p, c_int, c_void_p,
                                          c_int, c_void_p]
sqlite3_lib.sqlite3_bind_blob.restype = c_int
sqlite3_bind_blob = sqlite3_lib.sqlite3_bind_blob

sqlite3_lib.sqlite3_bind_double.argtypes = [sqlite3_stmt_p, c_int, c_double]
sqlite3_lib.sqlite3_bind_double.restype = c_int
sqlite3_bind_double = sqlite3_lib.sqlite3_bind_double

sqlite3_bind_int = sqlite3_lib.sqlite3_bind_int
sqlite3_lib.sqlite3_bind_int.argtypes = [sqlite3_stmt_p, c_int, c_int]
sqlite3_lib.sqlite3_bind_int.restype = c_int

sqlite3_lib.sqlite3_bind_null.argtypes = [sqlite3_stmt_p, c_int]
sqlite3_lib.sqlite3_bind_null.restype = c_int
sqlite3_bind_null = sqlite3_lib.sqlite3_bind_null

sqlite3_lib.sqlite3_bind_text.argtypes = [sqlite3_stmt_p, c_int, c_char_p,
                                          c_int, c_void_p]
sqlite3_lib.sqlite3_bind_text.restype = c_int
sqlite3_bind_text = sqlite3_lib.sqlite3_bind_text

# https://www.sqlite.org/c3ref/step.html
sqlite3_lib.sqlite3_step.argtypes = [sqlite3_stmt_p]
sqlite3_lib.sqlite3_step.restype = c_int
sqlite3_step = sqlite3_lib.sqlite3_step

# https://www.sqlite.org/c3ref/column_blob.html
sqlite3_lib.sqlite3_column_blob.argtypes = [sqlite3_stmt_p, c_int]
sqlite3_lib.sqlite3_column_blob.restype = c_void_p
sqlite3_column_blob = sqlite3_lib.sqlite3_column_blob

sqlite3_lib.sqlite3_column_double.argtypes = [sqlite3_stmt_p, c_int]
sqlite3_lib.sqlite3_column_double.restype = c_double
sqlite3_column_double = sqlite3_lib.sqlite3_column_double

sqlite3_lib.sqlite3_column_int.argtypes = [sqlite3_stmt_p, c_int]
sqlite3_lib.sqlite3_column_int.restype = c_int
sqlite3_column_int = sqlite3_lib.sqlite3_column_int

sqlite3_lib.sqlite3_column_text.argtypes = [sqlite3_stmt_p, c_int]
sqlite3_lib.sqlite3_column_text.restype = c_char_p
sqlite3_column_text = sqlite3_lib.sqlite3_column_text

sqlite3_lib.sqlite3_column_bytes.argtypes = [sqlite3_stmt_p, c_int]
sqlite3_lib.sqlite3_column_bytes.restype = c_int
sqlite3_column_bytes = sqlite3_lib.sqlite3_column_bytes

sqlite3_lib.sqlite3_column_type.argtypes = [sqlite3_stmt_p, c_int]
sqlite3_lib.sqlite3_column_type.restype = c_int
sqlite3_column_type = sqlite3_lib.sqlite3_column_type

# https://www.sqlite.org/c3ref/finalize.html
sqlite3_lib.sqlite3_finalize.argtypes = [sqlite3_stmt_p]
sqlite3_lib.sqlite3_finalize.restype = c_int
sqlite3_finalize = sqlite3_lib.sqlite3_finalize

# https://www.sqlite.org/c3ref/column_name.html
sqlite3_lib.sqlite3_column_name.argtypes = [sqlite3_stmt_p, c_int]
sqlite3_lib.sqlite3_column_name.restype = c_char_p
sqlite3_column_name = sqlite3_lib.sqlite3_column_name

# https://www.sqlite.org/c3ref/reset.html
sqlite3_lib.sqlite3_reset.argtypes = [sqlite3_stmt_p]
sqlite3_lib.sqlite3_reset.restype = c_int
sqlite3_reset = sqlite3_lib.sqlite3_reset

# https://www.sqlite.org/c3ref/clear_bindings.html
sqlite3_lib.sqlite3_clear_bindings.argtypes = [sqlite3_stmt_p]
sqlite3_lib.sqlite3_clear_bindings.restype = c_int
sqlite3_clear_bindings = sqlite3_lib.sqlite3_clear_bindings

# http://www.sqlite.org/c3ref/changes.html
sqlite3_lib.sqlite3_changes.argtypes = [sqlite3_p]
sqlite3_lib.sqlite3_changes.restype = c_int
sqlite3_changes = sqlite3_lib.sqlite3_changes

##############################################
# https://www.sqlite.org/c3ref/constlist.html
##############################################
# https://www.sqlite.org/c3ref/c_open_autoproxy.html
SQLITE_OPEN_READONLY = 0x00000001  # Ok for sqlite3_open_v2()
SQLITE_OPEN_READWRITE = 0x00000002  # Ok for sqlite3_open_v2()
SQLITE_OPEN_CREATE = 0x00000004  # Ok for sqlite3_open_v2()

# https://www.sqlite.org/rescode.html
SQLITE_ABORT = 4
SQLITE_AUTH = 23
SQLITE_BUSY = 5
SQLITE_CANTOPEN = 14
SQLITE_CONSTRAINT = 19
SQLITE_CORRUPT = 11
SQLITE_DONE = 101
SQLITE_EMPTY = 16
SQLITE_ERROR = 1
SQLITE_FORMAT = 24
SQLITE_FULL = 13
SQLITE_INTERNAL = 2
SQLITE_INTERRUPT = 9
SQLITE_IOERR = 10
SQLITE_LOCKED = 6
SQLITE_MISMATCH = 20
SQLITE_MISUSE = 21
SQLITE_NOLFS = 22
SQLITE_NOMEM = 7
SQLITE_NOTADB = 26
SQLITE_NOTFOUND = 12
SQLITE_NOTICE = 27
SQLITE_OK = 0
SQLITE_PERM = 3
SQLITE_PROTOCOL = 15
SQLITE_RANGE = 25
SQLITE_READONLY = 8
SQLITE_ROW = 100
SQLITE_SCHEMA = 17
SQLITE_TOOBIG = 18
SQLITE_WARNING = 28

# https://www.sqlite.org/c3ref/c_blob.html
SQLITE_INTEGER = 1
SQLITE_FLOAT = 2
SQLITE_TEXT = 3
SQLITE_BLOB = 4
SQLITE_NULL = 5

# https://www.sqlite.org/c3ref/c_static.html
SQLITE_STATIC = 0
SQLITE_TRANSIENT = -1


# types
class SQLType(enum.Enum):
    """
    List of SQLite types that can be bound.
    Values are internal to py4lo_sqlite and are currently not used.
    """
    TYPE_BLOB = 1
    TYPE_BLOB64 = 2
    TYPE_DOUBLE = 3
    TYPE_INT = 4
    TYPE_INT64 = 5
    NULL = 6
    TYPE_TEXT = 7
    TYPE_TEXT16 = 8
    TYPE_TEXT64 = 9
    TYPE_VALUE = 10
    TYPE_POINTER = 11
    TYPE_ZEROBLOB = 12
    TYPE_ZEROBLOB64 = 13


class Sqlite3Statement:
    """
    A SQLite3 statement. See `Sqlite3Database.prepare`.
    """
    def __init__(self, db: sqlite3_p, stmt: sqlite3_stmt_p):
        """
        @param db: sqlite3 handle (see https://www.sqlite.org/c3ref/open.html).
        @param stmt: the sqlite3 statement (see
         https://www.sqlite.org/c3ref/prepare.html)
        """
        self._db = db
        self._stmt = stmt

    def bind_text(self, i: int, v: str):
        """
        Bind a text parameter

        @param i: number of the col
        @param v: the text value
        """
        bs = v.encode("utf-8")
        ret = sqlite3_bind_text(self._stmt, i, bs, len(bs), SQLITE_TRANSIENT)
        if ret != SQLITE_OK:
            raise self._err(ret)

    def bind_blob(self, i: int, v: bytes):
        """
        Bind a blob parameter

        @param i: number of the col
        @param v: the blob value
        """
        ret = sqlite3_bind_blob(self._stmt, i, v, len(v), SQLITE_TRANSIENT)
        if ret != SQLITE_OK:
            raise self._err(ret)

    def bind_double(self, i: int, v: float):
        """
        Bind a double parameter (float in Python)

        @param i: number of the col
        @param v: the float value
        """
        ret = sqlite3_bind_double(self._stmt, i, v)
        if ret != SQLITE_OK:
            raise self._err(ret)

    def bind_int(self, i: int, v: int):
        """
        Bind an integer parameter

        @param i: number of the col
        @param v: the int value
        """
        ret = sqlite3_bind_int(self._stmt, i, v)
        if ret != SQLITE_OK:
            raise self._err(ret)

    def bind_null(self, i: int):
        """
        Bind a parameter to a null value.

        @param i: number of the col
        """
        ret = sqlite3_bind_null(self._stmt, i)
        if ret != SQLITE_OK:
            raise self._err(ret)

    def _err(self, ret: int) -> SQLiteError:
        return SQLiteError(ret, sqlite3_errmsg(self._db).decode("utf-8"))

    def reset(self):
        """
        Reset the statement but keep the current bindings
        (see https://www.sqlite.org/c3ref/reset.html)
        """
        ret = sqlite3_reset(self._stmt)
        if ret != SQLITE_OK:
            raise self._err(ret)

    def clear_bindings(self):
        """
        Reset all bindings (see https://www.sqlite.org/c3ref/reset.html)
        """
        ret = sqlite3_clear_bindings(self._stmt)
        if ret != SQLITE_OK:
            raise self._err(ret)

    def execute_update(self) -> int:
        """
        See https://www.sqlite.org/c3ref/step.html
        @return: the number of rows modified, inserted or deleted
        (see https://www.sqlite.org/c3ref/changes.html)
        """
        ret = sqlite3_step(self._stmt)
        if ret != SQLITE_DONE:
            raise self._err(ret)
        return sqlite3_changes(self._db)

    def execute_query(self, with_names: bool = False
                      ) -> Iterator[Union[List[Any], Mapping[str, Any]]]:
        """
        Execute the query.

        @param with_names: use the name for the return
        @return: an iterator over results. if `with_names`is True,
        return a mapping col name -> value, else return a list of values
        """
        if with_names:
            return self._execute_query_with_names()
        else:
            return self._execute_query_without_names()

    def _execute_query_with_names(self) -> Iterator[Mapping[str, Any]]:
        col_count = sqlite3_column_count(self._stmt)
        ret = sqlite3_step(self._stmt)
        names = [
            sqlite3_column_name(self._stmt, i).decode("utf-8")
            for i in range(col_count)
        ]
        column_types = [
            sqlite3_column_type(self._stmt, i)
            for i in range(col_count)
        ]

        while ret == SQLITE_ROW:
            row = dict([
                (names[i], self._value(i, column_types))
                for i in range(col_count)
            ])
            yield row
            ret = sqlite3_step(self._stmt)
        if ret != SQLITE_DONE:
            raise self._err(ret)

    def _execute_query_without_names(self) -> Iterator[List[Any]]:
        col_count = sqlite3_column_count(self._stmt)
        ret = sqlite3_step(self._stmt)
        column_types = [
            sqlite3_column_type(self._stmt, i)
            for i in range(col_count)
        ]

        while ret == SQLITE_ROW:
            row = [
                self._value(i, column_types)
                for i in range(col_count)
            ]
            yield row
            ret = sqlite3_step(self._stmt)
        if ret != SQLITE_DONE:
            raise self._err(ret)

    def _value(self, i: int, column_types: List[int]) -> Any:
        sql_type = column_types[i]
        if sql_type == SQLITE_INTEGER:
            ret = sqlite3_column_int(self._stmt, i)
        elif sql_type == SQLITE_FLOAT:
            ret = sqlite3_column_double(self._stmt, i)
        elif sql_type == SQLITE_TEXT:
            ret = sqlite3_column_text(self._stmt, i).decode("utf-8")
        elif sql_type == SQLITE_BLOB:
            ptr = sqlite3_column_blob(self._stmt, i)
            size = sqlite3_column_bytes(self._stmt, i)
            ret = string_at(ptr, size)
        elif sql_type == SQLITE_NULL:
            ret = None
        else:
            raise ValueError("Unknown type {}".format(sql_type))
        return ret


class Sqlite3Database:
    """
    A wrapper for a sqlite3 handle (see https://www.sqlite.org/c3ref/open.html).
    """
    def __init__(self, db: sqlite3_p):
        """
        @param db: SQLite db handle
        """
        self._db = db

    def execute_update(self, sql: str) -> int:
        """
        Execute an update (see https://www.sqlite.org/c3ref/exec.html).

        @param sql: the sql query
        @return: the number of rows modified, inserted or deleted
        (see https://www.sqlite.org/c3ref/changes.html)
        @raise SQLiteError: if the result is not OK
        """
        ret = sqlite3_exec(self._db, sql.encode("utf-8"), None, None, None)
        if ret != SQLITE_OK:
            raise self._err(ret)
        return sqlite3_changes(self._db)

    def _err(self, ret: int) -> SQLiteError:
        return SQLiteError(ret, sqlite3_errmsg(self._db).decode("utf-8"))

    @contextmanager
    def prepare(self, sql: str) -> Generator[Sqlite3Statement, None, None]:
        """
        Context manager to prepare a SQL statement:
        ```
        with db.prepare("...") as stmt:
            ...
        ```

        @param sql: the statement
        @yield: the SQLite3Statement
        """
        stmt_p = sqlite3_stmt_p()
        ret = sqlite3_prepare_v2(self._db, sql.encode("utf-8"), -1,
                                 byref(stmt_p), None)
        if ret != SQLITE_OK:
            raise self._err(ret)

        try:
            yield Sqlite3Statement(self._db, stmt_p)
        except Exception:
            raise
        finally:
            ret = sqlite3_finalize(stmt_p)
            if ret != SQLITE_OK:
                raise self._err(ret)

    @contextmanager
    def transaction(self, mode: TransactionMode = TransactionMode.DEFERRED
                    ) -> Generator[None, None, None]:
        """
        Context manager for a transaction:
        ```
        with db.transaction():
            ... # use db here
        ```

        If there is an exception, then we have the following sequence:
        ```
        BEGIN <mode> TRANSACTION
        <SQL commands> -- the exception here
        ROLLBACK
        ```
        and yield the exception to Python.

        If there is no exception:
        ```
        BEGIN <mode> TRANSACTION
        <SQL commands> -- no exception
        END TRANSACTION
        ```

        @param mode: transaction mode
        @yield: a context
        """
        self.execute_update("BEGIN {} TRANSACTION".format(mode.value))
        try:
            yield
        except SQLiteError:
            self.execute_update("ROLLBACK")
            raise
        else:
            self.execute_update("END TRANSACTION") # synonym of COMMIT

    def interrupt(self) -> bool:
        sqlite3_interrupt(self._db)
        return bool(sqlite3_is_interrupted(self._db))


@contextmanager
def sqlite_open(
        filepath: Union[str, Path], mode: str = "r", timeout: int=-1
) -> Generator[Sqlite3Database, None, None]:
    """
    Open a SQLite database in a context manager:
    ```
    with sqlite_open(...) as db:
        ...
    ```

    @param filepath: the path to the file
    @param mode: the mode ("r" for read, "rw" for readwrite
    or "crw" for create read write)
    @param timeout: an optional timeout
    @yield: a Sqlite3Database object
    """
    db = c_void_p()

    if isinstance(filepath, str):
        str_path = filepath
        path = Path(filepath)
    else:
        str_path = str(filepath.absolute())
        path = filepath

    if mode == "r":
        flags = SQLITE_OPEN_READONLY
        if not path.exists():
            raise FileNotFoundError(str_path)
    elif mode == "rw":
        flags = SQLITE_OPEN_READWRITE
        if not path.exists():
            raise FileNotFoundError(str_path)
    elif mode == "crw":
        flags = SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE
    else:
        raise ValueError(mode)

    sqlite3_open_v2(str_path.encode("utf-8"), byref(db), flags, None)
    if timeout > 0:
        sqlite3_busy_timeout(db, timeout)
    try:
        yield Sqlite3Database(db)
    finally:
        sqlite3_close_v2(db)
