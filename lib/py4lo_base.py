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
# mypy: disable-error-code="import-not-found"
import logging
import time
from pathlib import Path
from typing import Iterable, Union, Dict, Mapping, cast

from py4lo_helper import (
    uno_path_to_url,
    create_uno_service,
    to_items,
    remove_all,
)
from py4lo_typing import (
    lazy,
    UnoStatement,
    UnoConnection,
    UnoDocumentDataSource,
    UnoContainer,
    UnoDatabaseContext,
    UnoTable,
)

try:

    class DataType:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sdbc.DataType import (
            BIT,
            TINYINT,
            SMALLINT,
            INTEGER,
            BIGINT,
            FLOAT,
            REAL,
            DOUBLE,
            NUMERIC,
            DECIMAL,
            CHAR,
            VARCHAR,
            LONGVARCHAR,
            DATE,
            TIME,
            TIMESTAMP,
            BINARY,
            VARBINARY,
            LONGVARBINARY,
            SQLNULL,
            OTHER,
            OBJECT,
            DISTINCT,
            STRUCT,
            ARRAY,
            BLOB,
            CLOB,
            REF,
            BOOLEAN,
        )

except (ModuleNotFoundError, ImportError):
    from mock_constants import (  # type:ignore[assignment]
        DataType,
    )


class BaseTableBuilder:
    """
    A builder for a base table
    """

    _logger = logging.getLogger(__name__)

    def __init__(self, oConnection: UnoConnection, name: str):
        """
        @param oConnection: the connection
        @param name: the name of the table
        """
        self._oConnection = oConnection
        self._oTables = oConnection.Tables
        self._oTableDescriptor = self._oTables.createDataDescriptor()
        self._oTableDescriptor.Name = name
        self._oCols = self._oTableDescriptor.getColumns()

    def build(self):
        """
        Build the table
        """
        self._oTables.appendByDescriptor(self._oTableDescriptor)
        self._oConnection.commit()

    def add_column(self, col_name: str, col_type: DataType, **kwargs):
        """
        :param col_name: the name of the column
        :param col_type: the type of the column
        :param kwargs: Precision, ...
        """
        oCol = self._oCols.createDataDescriptor()
        oCol.Name = col_name
        oCol.Type = col_type
        for k, v in kwargs.items():
            setattr(oCol, k, v)
        self._oCols.appendByDescriptor(oCol)


def drop_all(oDrop: UnoContainer):
    """
    Drop all elements (count or name)
    @param oDrop: the container
    """
    try:
        for name in oDrop.ElementNames:
            oDrop.dropByName(name)
    except AttributeError:
        while oDrop.Count:
            oDrop.dropByIndex(0)


class BaseDB:
    """
    A Database wrapper.
    """

    _logger = logging.getLogger(__name__)

    def __init__(self, oDB: UnoDocumentDataSource):
        self._oDB = oDB
        self._oConnection = lazy(UnoConnection)
        self._oStatement = lazy(UnoStatement)
        self._sql_by_name = cast(Dict[str, str], {})

    @property
    def connection(self) -> UnoConnection:
        if self._oConnection is None or self._oConnection.isClosed():
            oHandler = create_uno_service(
                "com.sun.star.sdb.InteractionHandler"
            )
            self._oConnection = self._oDB.connectWithCompletion(oHandler)

        return self._oConnection

    @property
    def statement(self) -> UnoStatement:
        if self._oStatement is None:
            self._oStatement = self.connection.createStatement()
        return self._oStatement

    def has_table(self, name: str) -> bool:
        return self.get_tables().hasByName(name)

    def get_table_builder(self, name: str) -> BaseTableBuilder:
        oConnection = self.connection
        oTables = oConnection.Tables
        if oTables.hasByName(name):
            raise ValueError("Exists")

        return BaseTableBuilder(oConnection, name)

    def get_tables(self) -> UnoContainer:
        return self.connection.Tables

    def execute_update(self, sql: str):
        self.statement.executeUpdate(sql)

    def execute(self, sql: str):
        self.statement.execute(sql)

    def commit(self):
        self.connection.commit()

    def add_batch(self, sql: str):
        self.statement.addBatch(sql)

    def execute_batch(self):
        self.statement.executeBatch()

    def save(self):
        self.connection.commit()
        self._oDB.DatabaseDocument.store()

    def close(self):
        self.save()
        self.connection.close()

    def remove_table(self, name: str):
        self.get_tables().removeByName(name)

    def add_pk(self, table: str, field_s: Union[str, Iterable[str]]):
        """
        Add a primary key
        @param table: the table name
        @param field_s: the field or the fields
        """
        if isinstance(field_s, str):
            fields = [field_s]
        else:
            fields = list(field_s)
        sql_lines = (
            ["ALTER TABLE {}".format(table)]
            + ["    ALTER {} SET NOT NULL,".format(field) for field in fields]
            + [
                "    ADD CONSTRAINT PK_{} PRIMARY KEY ({})".format(
                    table, ", ".join(fields)
                )
            ]
        )

        self.execute("\n".join(sql_lines))

    def create_idx(self, table: str, field_s: Union[str, Iterable[str]]):
        """
        Add an index
        @param table: the table name
        @param field_s: the field or the fields
        """
        if isinstance(field_s, str):
            fields = [field_s]
        else:
            fields = list(field_s)
        sql = 'CREATE INDEX IDX_{}_{} ON {} ("{}")'.format(
            table, "_".join(fields), table, '", "'.join(fields)
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

    def drop_tables(self):
        """
        Drop all tables from the connection. Try 3 times.
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
                break

    def _drop_table(self, oTables: UnoContainer[UnoTable], name: str):
        oTable = oTables.getByName(name)
        drop_all(oTable.Keys)
        drop_all(oTable.Indexes)
        oTables.dropByName(name)

    def remove_queries(self):
        remove_all(self._oDB.QueryDefinitions)

    def secure_drop_table(self, name: str):
        oTables = self.get_tables()
        if not oTables.hasByName(name):
            return

        self._drop_table(oTables, name)

    def get_queries(self) -> Dict[str, str]:
        """
        @return: a mapping name -> sql.
        """
        sql_by_name = {}
        for name in self._oDB.QueryDefinitions.ElementNames:
            oQuery = self._oDB.QueryDefinitions.getByName(name)
            sql_by_name[name] = oQuery.Command
        return sql_by_name

    def add_queries(self, sql_by_name: Mapping[str, str]):
        """
        Add some queries
        """
        for name, sql in sql_by_name.items():
            self.add_query(name, sql)

    def add_query(self, name: str, sql: str):
        """
        Add a query
        """
        oQuery = self._oDB.QueryDefinitions.createInstance()
        oQuery.Command = sql
        # noinspection PyBroadException
        try:
            self._oDB.QueryDefinitions.insertByName(name, oQuery)
        except Exception:
            self._logger.exception(
                "insert query %s => %s", repr(name), repr(sql)
            )


def open_or_create_db(path: Path) -> "BaseDB":
    oDBContext = cast(
        UnoDatabaseContext,
        create_uno_service("com.sun.star.sdb.DatabaseContext"),
    )
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
