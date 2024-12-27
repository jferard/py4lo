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
# mypy: disable-error-code="import-not-found"
"""
A set of command to pilot a LibreOffice Base document.
"""

import logging
import time
from pathlib import Path
from typing import Iterable, Union, Dict, Mapping

from py4lo_helper import (uno_path_to_url, create_uno_service, to_items,
                          remove_all)
from py4lo_typing import UnoObject, UnoService, lazy

try:
    class DataType:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sdbc.DataType import (
            BIT, TINYINT, SMALLINT, INTEGER, BIGINT, FLOAT, REAL, DOUBLE,
            NUMERIC, DECIMAL, CHAR, VARCHAR, LONGVARCHAR, DATE, TIME,
            TIMESTAMP, BINARY, VARBINARY, LONGVARBINARY, SQLNULL, OTHER,
            OBJECT, DISTINCT, STRUCT, ARRAY, BLOB, CLOB, REF, BOOLEAN
        )

    class ColumnValue:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sdbc.ColumnValue import (
            NO_NULLS, NULLABLE, NULLABLE_UNKNOWN
        )

except (ModuleNotFoundError, ImportError):
    from _mock_constants import ( # type:ignore[assignment]
        DataType,
        ColumnValue  # noqa: F401
    )


class BaseTableBuilder:
    """
    `BaseTableBuilder` is a builder for a database table. Use with a
    `BaseDB` object.
    Example :

    ```
    builder = base_db.get_table_builder("persons")
    builder.add_column("person_id", DataType.INT)
    builder.add_column("name", DataType.CHAR, Precision=100)
    builder.build()
    ```
    """
    _logger = logging.getLogger(__name__)

    def __init__(self, oConnection: UnoObject, name: str):
        """
        @param oConnection: the connection (a com.sun.star.sdbc.XConnection object)
        @param name: the name of the table to build
        """
        self._oConnection = oConnection
        self._oTables = oConnection.Tables
        self._oTableDescriptor = self._oTables.createDataDescriptor()
        self._oTableDescriptor.Name = name
        self._oCols = self._oTableDescriptor.Columns

    def build(self):
        """
        Build the table and add it to the database.
        """
        self._oTables.appendByDescriptor(self._oTableDescriptor)
        self._oConnection.commit()

    def add_column(self, col_name: str, col_type: DataType, **kwargs):
        """
        Add a column to the future table.

        :param col_name: the name of the column
        :param col_type: the type of the column (see: com.sun.star.sdbc.DataType)
        :param kwargs: Precision, ... (see: com.sun.star.sdbcx.ColumnDescriptor)
        """
        oCol = self._oCols.createDataDescriptor()
        oCol.Name = col_name
        oCol.Type = col_type
        for k, v in kwargs.items():
            setattr(oCol, k, v)
        self._oCols.appendByDescriptor(oCol)


def drop_all(oDrop: UnoObject):
    """
    Drop all elements (count or name) of a container. The container might be
    the table views, indexes... or the database tables for instance.

    @param oDrop: the container (see: com.sun.star.sdbcx.XDrop)
    """
    try:
        for name in oDrop.ElementNames:
            oDrop.dropByName(name)
    except AttributeError:
        while oDrop.Count:
            oDrop.dropByIndex(0)


class BaseDB:
    """
    A batabase wrapper.
    """
    _logger = logging.getLogger(__name__)

    def __init__(self, oDB: UnoObject):
        """
        @param oDB: a database context (see: com.sun.star.sdb.DatabaseContext)
        """
        self._oDB = oDB
        self._oConnection = lazy(UnoService)
        self._oStatement = lazy(UnoService)
        self._sql_by_name = {}

    @property
    def connection(self) -> UnoObject:
        """
        Returns the current connection or opens a new one.

        @return: a database connection (see: com.sun.star.sdb.XConnection)
        """
        if self._oConnection is None or self._oConnection.isClosed():
            oHandler = create_uno_service(
                "com.sun.star.sdb.InteractionHandler")
            self._oConnection = self._oDB.connectWithCompletion(oHandler)

        return self._oConnection

    @property
    def statement(self) -> UnoObject:
        """
        Returns the current statement or opens a new one.

        @return: a statement (see: com.sun.star.sdb.XStatement)
        """
        if self._oStatement is None:
            self._oStatement = self.connection.createStatement()
        return self._oStatement

    def has_table(self, name: str) -> bool:
        """
        @param name: a table name
        @return: True if a table having this name exists, False otherwise.
        """
        return self.get_tables().hasByName(name)

    def get_table_builder(self, name: str) -> BaseTableBuilder:
        """
        @param name: the future table name
        @return: the builder for the table.
        @raises: ValueError if a table having this name exists
        """
        oConnection = self.connection
        oTables = oConnection.Tables
        if oTables.hasByName(name):
            raise ValueError("Exists")

        return BaseTableBuilder(oConnection, name)

    def get_tables(self) -> UnoObject:
        """
        @return: the table container (a XNameAccess)
        """
        return self.connection.Tables

    def execute_update(self, sql: str):
        """
        Execute an database update query (see: com.sun.star.sdb.XConnection.executeUpdate)

        @param sql: the query to execute (DDL, DML or DCL query)
        """
        self.statement.executeUpdate(sql)

    def execute(self, sql: str):
        """
        Execute an database SELECT query (see: com.sun.star.sdb.XConnection.execute)

        @param sql: the query to excute
        @return: a result set (see: com.sun.star.sdb.XResultSet)
        """
        self.statement.execute(sql)

    def commit(self):
        """
        Make all changes since previous commit/rollback. (see: com.sun.star.sdb.XConnection.commit)
        """
        self.connection.commit()

    def add_batch(self, sql: str):
        """
        Add a sql command to the batch of commands (see: com.sun.star.sdb.XBatchExecution.addBatch).

        @param sql: the sql command
        """
        self.statement.addBatch(sql)

    def execute_batch(self):
        """
        Execute the batch of commands (see: com.sun.star.sdb.XBatchExecution.executeBatch).
        """
        self.statement.executeBatch()

    def save(self):
        """
        Commit the changes and save the database
        """
        self.connection.commit()
        self._oDB.DatabaseDocument.store()

    def close(self):
        """
        Commit the changes, save the database and closes the connection
        """
        self.save()
        self.connection.close()

    def remove_table(self, name: str):
        """
        Remove a table
        @param name: the name of the table to remove
        """
        self.get_tables().removeByName(name)

    def add_pk(self, table: str, field_s: Union[str, Iterable[str]]):
        """
        Add a primary key to a table.
        @param table: the table name
        @param field_s: the field or the fields to use
        """
        if isinstance(field_s, str):
            fields = [field_s]
        else:
            fields = list(field_s)
        sql_lines = ["ALTER TABLE {}".format(table)] + [
            "    ALTER {} SET NOT NULL,".format(field) for field in fields
        ] + ["    ADD CONSTRAINT PK_{} PRIMARY KEY ({})".format(
            table, ", ".join(fields))]

        self.execute("\n".join(sql_lines))

    def create_idx(self, table: str, field_s: Union[str, Iterable[str]]):
        """
        Add an index to table
        @param table: the table name
        @param field_s: the field or the fields
        """
        if isinstance(field_s, str):
            fields = [field_s]
        else:
            fields = list(field_s)

        idx_name = "IDX_{}_{}".format(table, "_".join(fields))
        idx_name = idx_name[:31]  # Firebird limit
        sql = "CREATE INDEX {} ON {} (\"{}\")".format(
            idx_name, table, "\", \"".join(fields)
        )
        # noinspection PyBroadException
        try:
            self.execute(sql)
        except Exception:
            self._logger.exception(sql)

    def drop_views(self):
        """
        Drop all views from the connection
        """
        drop_all(self.connection.Views)

    def drop_tables(self) -> bool:
        """
        Drop all tables from the database. Try 3 times and then give up.
        @return: True if the table where removed, False otherwise.
        """
        for i in range(3):
            # noinspection PyBroadException
            try:
                oTables = self.get_tables()
                for name, oTable in to_items(oTables):
                    self._drop_table(oTables, name)
            except Exception:
                time.sleep(1)
            else:
                return True
        return False

    def _drop_table(self, oTables: UnoObject, name: str):
        oTable = oTables.getByName(name)
        drop_all(oTable.Keys)
        drop_all(oTable.Indexes)
        oTables.dropByName(name)

    def remove_queries(self):
        """
        Remove all queries from the data source
        """
        remove_all(self._oDB.QueryDefinitions)

    def secure_drop_table(self, name: str) -> bool:
        """
        Remove a table from the database.

        @param name: the name of the table to remove
        @return: True if the table was removed, False otherwise.
        """
        oTables = self.get_tables()
        try:
            self._drop_table(oTables, name)
        except Exception:
            return False
        else:
            return True

    def get_queries(self) -> Dict[str, str]:
        """
        Returns all the queries from a data source.

        @return: a mapping name -> sql query.
        """
        sql_by_name = {}
        for name in self._oDB.QueryDefinitions.ElementNames:
            oQuery = self._oDB.QueryDefinitions.getByName(name)
            sql_by_name[name] = oQuery.Command
        return sql_by_name

    def add_queries(self, sql_by_name: Mapping[str, str]):
        """
        Add some queries (by name)
        :@param sql_by_name: a mapping name -> sql query.
        """
        for name, sql in sql_by_name.items():
            self.add_query(name, sql)

    def add_query(self, name: str, sql: str):
        """
        Add a query to the data source.

        @param name: the name of the query
        @param sql: the sql query.
        """
        oQuery = self._oDB.QueryDefinitions.createInstance()
        oQuery.Command = sql
        # noinspection PyBroadException
        try:
            self._oDB.QueryDefinitions.insertByName(name, oQuery)
        except Exception:
            self._logger.exception("insert query %s => %s", repr(name),
                                   repr(sql))


def open_or_create_db(path: Path) -> "BaseDB":
    """
    Open or create a data base (actually a data source).
    Example

    ```
    base_db = open_or_create(Path(r"./mybase.odb")
    ```

    If created, the database is immediately saved.

    @param path: the path
    @return: the `BaseDB` wrapper
    """
    oDBContext = create_uno_service("com.sun.star.sdb.DatabaseContext")
    url = uno_path_to_url(path)
    # noinspection PyBroadException
    try:
        oDB = oDBContext.getByName(url)
    except Exception:
        oDB = oDBContext.createInstance()
        oDB.URL = "sdbc:embedded:firebird"
        oDocument = oDB.DatabaseDocument
        oDocument.storeAsURL(url, tuple())

    return BaseDB(oDB)
