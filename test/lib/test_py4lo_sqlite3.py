import ctypes
import random
import string
import threading
import unittest
from datetime import datetime
from time import sleep
from pathlib import Path
from unittest import mock

from py4lo_sqlite3 import (
    sqlite_open, SQLiteError, TransactionMode, SQLITE_BUSY, SQLITE_ERROR,
    SQLITE_CONSTRAINT, Sqlite3Database, SQLITE_OK
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
                "CREATE TABLE t(a INTEGER, b TEXT, oTextRange REAL, e BLOB)"))

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

        with self.assertRaises(SQLiteError) as cm:
            with sqlite_open(self._path, "rw") as db, sqlite_open(self._path, "rw") as db2:
                with db.transaction(TransactionMode.IMMEDIATE):
                    db2.execute_update("INSERT INTO t VALUES (1)")

        exc = cm.exception
        self.assertEqual(SQLITE_BUSY, exc.result_code)
        self.assertEqual('database is locked', exc.msg)

    def test_timeout_long(self):
        with sqlite_open(self._path, "crw") as db:
            db.execute_update("CREATE TABLE t(x INTEGER)")

        with sqlite_open(self._path, "rw") as db, sqlite_open(self._path, "rw", 2_000) as db2:
            def func():
                with db.transaction(TransactionMode.IMMEDIATE):
                    db.execute_update("INSERT INTO t VALUES (1)")
                    sleep(1)

            t = threading.Thread(target=func)
            t.start()
            db2.execute_update("INSERT INTO t VALUES (2)")
            t.join()

    def test_timeout_short(self):
        with sqlite_open(self._path, "crw") as db:
            db.execute_update("CREATE TABLE t(x INTEGER)")

        with self.assertRaises(SQLiteError) as cm:
            with sqlite_open(self._path, "rw") as db, sqlite_open(self._path, "rw", 100) as db2:
                def func():
                    try:
                        with db.transaction(TransactionMode.IMMEDIATE):
                            db.execute_update("INSERT INTO t VALUES (1)")
                            sleep(1)
                    except Exception:
                        pass

                t = threading.Thread(target=func)
                t.start()
                db2.execute_update("INSERT INTO t VALUES (2)")
                t.join()

        exc = cm.exception
        self.assertEqual(SQLITE_BUSY, exc.result_code)
        self.assertEqual('database is locked', exc.msg)

    def test_bindings(self):
        with sqlite_open(self._path, "crw") as db:
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a INTEGER, b TEXT, oTextRange REAL, e BLOB) STRICT"))
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
                "CREATE TABLE t(a INTEGER, b TEXT, oTextRange REAL, e BLOB) STRICT"))
            self.assertEqual(0, db.execute_update(
                "CREATE UNIQUE INDEX `id_UNIQUE` ON `t` (`a` ASC)"))

    @mock.patch("py4lo_sqlite3.sqlite3_errmsg")
    @mock.patch("py4lo_sqlite3.sqlite3_changes")
    @mock.patch("py4lo_sqlite3.sqlite3_exec")
    def test_rollback(self, sqlite3_exec, sqlite3_changes, sqlite3_errmsg):
        # Arrange
        sqlite3_exec.side_effect = [SQLITE_OK, SQLITE_ERROR, SQLITE_OK]
        sqlite3_changes.side_effect = [0, 0, 0]
        sqlite3_errmsg.side_effect = [b"Err"]

        db = mock.Mock()
        sqlite3_db = Sqlite3Database(db)

        # Act
        with self.assertRaises(SQLiteError):
            with sqlite3_db.transaction(TransactionMode.IMMEDIATE):
                sqlite3_db.execute_update("FOO")

        # Assert
        self.assertEqual(
            [
                mock.call(db, b'BEGIN IMMEDIATE TRANSACTION', None, None, None),
                mock.call(db, b'FOO', None, None, None),
                mock.call(db, b'ROLLBACK', None, None, None),
            ],
            sqlite3_exec.mock_calls
        )

    @mock.patch("py4lo_sqlite3.sqlite3_changes")
    @mock.patch("py4lo_sqlite3.sqlite3_exec")
    def test_commit(self, sqlite3_exec, sqlite3_changes):
        # Arrange
        sqlite3_exec.side_effect = [SQLITE_OK, SQLITE_OK, SQLITE_OK]
        sqlite3_changes.side_effect = [0, 1, 0]

        db = mock.Mock()
        sqlite3_db = Sqlite3Database(db)

        # Act
        with sqlite3_db.transaction(TransactionMode.IMMEDIATE):
            sqlite3_db.execute_update("INSERT INTO t VALUES (1)")

        # Assert
        self.assertEqual(
            [
                mock.call(db, b'BEGIN IMMEDIATE TRANSACTION', None, None, None),
                mock.call(db, b'INSERT INTO t VALUES (1)', None, None, None),
                mock.call(db, b'END TRANSACTION', None, None, None),
            ],
            sqlite3_exec.mock_calls
        )


if __name__ == '__main__':
    unittest.main()
