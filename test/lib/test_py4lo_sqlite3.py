import ctypes
import random
import string
import unittest
from datetime import datetime
from pathlib import Path

from py4lo_sqlite3 import (
    sqlite_open, SQLiteError, TransactionMode, SQLITE_BUSY, SQLITE_ERROR,
    SQLITE_CONSTRAINT
)


def randbytes(n):
    try:  # >= 3.9
        return random.randbytes(random.randrange(0, n))
    except AttributeError:  # <= 3.8
        return bytes([random.randrange(0, 256) for _ in range(0, n)])


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
            n = 1000
            data = [
                (
                    random.randint(0, 100),
                    "".join(random.choice(string.ascii_letters) for _ in
                            range(random.randrange(10, 200))),
                    random.random() * 100,
                    randbytes(n)
                ) for _ in range(10000)
            ]

            t2 = datetime.now()
            print(t2 - t1)
            t1 = t2

            print("-> create table")
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a INTEGER, b TEXT, c REAL, e BLOB)"))

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
        self.assertEqual(SQLITE_ERROR, exc.result_code)
        self.assertEqual('near "Hello": syntax error', exc.msg)

    def test_busy(self):
        with sqlite_open(self._path, "crw") as db:
            db.execute_update("CREATE TABLE t(x INTEGER)")

        with sqlite_open(self._path, "rw") as db, sqlite_open(self._path,
                                                              "rw") as db2:
            with db.transaction(TransactionMode.IMMEDIATE):
                try:
                    with db2.transaction(TransactionMode.IMMEDIATE):
                        pass
                except SQLiteError as exc:
                    self.assertEqual(SQLITE_BUSY, exc.result_code)
                    self.assertEqual('database is locked', exc.msg)

    def test_bindings(self):
        with sqlite_open(self._path, "crw") as db:
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a INTEGER, b TEXT, c REAL, e BLOB) STRICT"))
            with self.assertRaises(SQLiteError) as cm:
                with db.prepare("INSERT INTO t VALUES(?, ?, ?, ?)") as stmt:
                    with self.assertRaises(ctypes.ArgumentError):
                        stmt.bind_int(1, "ok")

                    stmt.bind_text(1, "ok")
                    stmt.execute_update()

                exc = cm.exception
                self.assertEqual(SQLITE_CONSTRAINT, exc.result_code)
                self.assertEqual(
                    'cannot store TEXT value in INTEGER column t.a', exc.msg)

    def test_index(self):
        with sqlite_open(self._path, "crw") as db:
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a INTEGER, b TEXT, c REAL, e BLOB) STRICT"))
            self.assertEqual(0, db.execute_update(
                "CREATE UNIQUE INDEX `id_UNIQUE` ON `t` (`a` ASC)"))


if __name__ == '__main__':
    unittest.main()
