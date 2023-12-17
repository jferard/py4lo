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
# py4lo: entry
# py4lo: embed lib py4lo_typing
# py4lo: embed lib py4lo_helper
# py4lo: embed lib py4lo_commons
# py4lo: embed lib py4lo_sqlite3
import logging
from pathlib import Path

from py4lo_helper import provider
from py4lo_commons import Commons, uno_url_to_path
from py4lo_sqlite3 import sqlite_open

try:
    assert provider is not None
    doc_path = uno_url_to_path(provider.doc.URL)
    assert doc_path is not None
    CUR_PATH = doc_path.parent
except (AssertionError, AttributeError):
    CUR_PATH = Path.cwd()

try:
    commons = Commons.create(XSCRIPTCONTEXT)  # type: ignore[name-defined]
except NameError:
    pass
else:
    logger = logging.getLogger()
    commons.init_logger(logger, CUR_PATH / "sqlite3.log")


def sqlite_example(*_args):
    path = CUR_PATH / "temp.sqlite3"
    logger.debug("Base path %s", path)
    try:
        oSheet = provider.controller.ActiveSheet
        data_array = oSheet.getCellRangeByName("table_rows").DataArray
        logger.debug("Table rows %s", data_array)
        query = oSheet.getCellRangeByName("query").String
        logger.debug("Query %s", query)

        path.unlink(True)

        with sqlite_open(path, "crw") as db:
            db.execute_update(
                "CREATE TABLE persons " "(name TEXT, age INTEGER, percentage DOUBLE)"
            )
            with db.transaction():
                with db.prepare("INSERT INTO persons VALUES(?, ?, ?)") as stmt:
                    for data_row in data_array:
                        stmt.reset()
                        stmt.clear_bindings()
                        stmt.bind_text(1, data_row[0])
                        stmt.bind_int(2, int(data_row[1]))
                        stmt.bind_double(3, data_row[2])
                        stmt.execute_update()

            logger.debug("Table full")
            with db.prepare(query) as stmt:
                new_data_array = list(stmt.execute_query())

        logger.debug("Query result %s", new_data_array)
        arr = list(oSheet.getCellRangeByName("output_rows").DataArray)
        arr = new_data_array[: len(new_data_array)] + arr[len(new_data_array) :]
        oSheet.getCellRangeByName("output_rows").DataArray = arr
    except Exception:
        logger.exception("EXC")
    finally:
        path.unlink(True)
