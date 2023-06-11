import random
import string
import time
import unittest
from datetime import datetime
from pathlib import Path
from pprint import pprint

from py4lo_sqlite3 import sqlite_open, SQLType, SQLiteError, TransactionMode


class Sqlite3TestCase(unittest.TestCase):
    def setUp(self) -> None:
        self._path = Path("test.sqlite3")
        self._path.unlink(True)

    def tearDown(self) -> None:
        self._path = Path("test.sqlite3")
        self._path.unlink(True)

    def test_sqlite3(self):
        with sqlite_open(self._path, "crw") as db:
            t1 = datetime.now()
            print("-> generate data")
            data = [
                (
                    random.randint(0, 100),
                    "".join(random.choice(string.ascii_letters) for _ in range(random.randrange(10, 200))),
                    random.random() * 100,
                    random.randbytes(random.randrange(0, 1000))
                ) for _ in range(10000)
            ]

            t2 = datetime.now()
            print(t2 - t1)
            t1 = t2

            print("-> create table")
            self.assertEqual(0, db.execute_update("CREATE TABLE t(a INTEGER, b TEXT, c DOUBLE, e BLOB)"))

            t2 = datetime.now()
            print(t2 - t1)
            t1 = t2

            print("-> fill table")
            with db.transaction():
                with db.prepare("INSERT INTO t VALUES(?, ?, ?, ?)") as stmt:
                    for data_row in data:
                        stmt.reset()
                        stmt.clear_bindings()
                        stmt.bind_int(1, data_row[0])
                        stmt.bind_text(2, data_row[1])
                        stmt.bind_double(3, data_row[2])
                        stmt.bind_blob(4, data_row[3])
                        try:
                            self.assertEqual(1, stmt.execute_update())
                        except Exception as e:
                            print(e)

            t2 = datetime.now()
            print(t2 - t1)
            t1 = t2

            print("-> select")
            with db.prepare("SELECT * FROM t") as stmt:
                for db_row, data_row in zip(stmt.execute_query(), data):
                    self.assertEqual(list(data_row), db_row)

            t2 = datetime.now()
            print(t2 - t1)

    def test_open_rw_missing_file(self):
        with self.assertRaises(FileNotFoundError):
            with sqlite_open(self._path, "rw") as db:
                print(db)

    def test_open_r_missing_file(self):
        with self.assertRaises(FileNotFoundError):
            with sqlite_open(self._path, "r") as db:
                print(db)

    def test_exec(self):
        with self.assertRaises(SQLiteError) as cm:
            with sqlite_open(self._path, "crw") as db:
                db.execute_update("Hello, world!")

        exc = cm.exception
        self.assertEqual(1, exc.result_code)
        self.assertEqual('near "Hello": syntax error', exc.msg)

    def test_busy(self):
        with sqlite_open(self._path, "crw") as db:
            db.execute_update("CREATE TABLE t(x INTEGER)")

        with sqlite_open(self._path, "rw") as db, sqlite_open(self._path, "rw") as db2:
            with db.transaction(TransactionMode.IMMEDIATE):
                try:
                    with db2.transaction(TransactionMode.IMMEDIATE):
                        pass
                except SQLiteError as exc:
                    self.assertEqual(5, exc.result_code)
                    self.assertEqual('database is locked', exc.msg)


if __name__ == '__main__':
    unittest.main()
