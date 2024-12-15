import ctypes
import random
import string
import threading
import unittest
from time import sleep
from pathlib import Path
from unittest import mock
import datetime as dt

from py4lo_sqlite3 import (
    sqlite_open, SQLiteError, TransactionMode, SQLITE_BUSY, SQLITE_ERROR,
    SQLITE_CONSTRAINT, Sqlite3Database, SQLITE_OK, SQLITE_TEXT, SQLITE_BLOB,
    SQLITE_INTEGER, SQLITE_FLOAT, decode_text_utf8_to_str, decode_blob_to_bytes,
    sqlite3_column_int, sqlite3_column_double, create_decode_text_to_str,
    datetime_to_julian, PY4LO_UNIX_TS, decode_unix_ts_to_datetime_utc,
    create_decode_unix_ts_to_datetime, PY4LO_JULIAN,
    decode_julian_to_datetime_utc, create_decode_julian_to_datetime,
    PY4LO_ISO8601, decode_iso8601_to_datetime
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
            t1 = dt.datetime.now()
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

            t2 = dt.datetime.now()
            print(t2 - t1)
            t1 = t2

            print("-> create table")
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a INTEGER, b TEXT, c REAL, d BLOB)"))

            t2 = dt.datetime.now()
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

            t2 = dt.datetime.now()
            print(t2 - t1)
            t1 = t2

            print("-> select")
            with db.prepare("SELECT * FROM t") as stmt:
                for db_row, data_row in zip(stmt.execute_query(), data):
                    self.assertEqual(list(data_row), db_row)

            t2 = dt.datetime.now()
            print(t2 - t1)

    def test_sqlite_with_names(self):
        with sqlite_open(self._path, "crw") as db:
            t1 = dt.datetime.now()
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

            t2 = dt.datetime.now()
            print(t2 - t1)
            t1 = t2

            print("-> create table")
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a INTEGER, b TEXT, c REAL, d BLOB)"))

            t2 = dt.datetime.now()
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

            t2 = dt.datetime.now()
            print(t2 - t1)
            t1 = t2

            print("-> select")
            with db.prepare("SELECT * FROM t") as stmt:
                for db_row, data_row in zip(stmt.execute_query(with_names=True), data):
                    self.assertEqual(dict(zip('abcd', data_row)), db_row)

            t2 = dt.datetime.now()
            print(t2 - t1)

    def test_text(self):
        with sqlite_open(self._path, "crw") as db:
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a TEXT)"))
            with db.transaction():
                with db.prepare("INSERT INTO t VALUES(?)") as stmt:
                    for value in (None, "foo", ""):
                        stmt.reset()
                        stmt.clear_bindings()
                        stmt.bind_text(1, value)
                        try:
                            self.assertEqual(1, stmt.execute_update())
                        except Exception as e:
                            print(e)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query())
                self.assertEqual([[None], ["foo"], [""]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[SQLITE_TEXT]))
                self.assertEqual([[None], ["foo"], [""]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[decode_text_utf8_to_str]))
                self.assertEqual([[None], ["foo"], [""]], db_rows)

    def test_blob(self):
        with sqlite_open(self._path, "crw") as db:
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a BLOB)"))
            with db.transaction():
                with db.prepare("INSERT INTO t VALUES(?)") as stmt:
                    for value in (None, b"foo", b"", b"F\xe9rard"):
                        stmt.reset()
                        stmt.clear_bindings()
                        stmt.bind_blob(1, value)
                        try:
                            self.assertEqual(1, stmt.execute_update())
                        except Exception as e:
                            print(e)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query())
                self.assertEqual([[None], [b"foo"], [b""], [b"F\xe9rard"]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[SQLITE_BLOB]))
                self.assertEqual([[None], [b"foo"], [b""], [b"F\xe9rard"]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[decode_blob_to_bytes]))
                self.assertEqual([[None], [b"foo"], [b""], [b"F\xe9rard"]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[create_decode_text_to_str("latin-1")]))
                self.assertEqual([[None], ["foo"], [""], ["Férard"]], db_rows)

    def test_blob_with_null_byte(self):
        with sqlite_open(self._path, "crw") as db:
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a BLOB)"))
            with db.transaction():
                with db.prepare("INSERT INTO t VALUES(?)") as stmt:
                    stmt.reset()
                    stmt.clear_bindings()
                    stmt.bind_blob(1, b"foo\0bar")
                    try:
                        self.assertEqual(1, stmt.execute_update())
                    except Exception as e:
                        print(e)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[SQLITE_BLOB]))
                self.assertEqual([[b"foo\0bar"]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[SQLITE_TEXT]))
                self.assertEqual([["foo"]], db_rows)


    def test_integer(self):
        with sqlite_open(self._path, "crw") as db:
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a INTEGER)"))
            with db.transaction():
                with db.prepare("INSERT INTO t VALUES(?)") as stmt:
                    for value in (None, 1, 0):
                        stmt.reset()
                        stmt.clear_bindings()
                        stmt.bind_int(1, value)
                        try:
                            self.assertEqual(1, stmt.execute_update())
                        except Exception as e:
                            print(e)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query())
                self.assertEqual([[None], [1], [0]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[SQLITE_INTEGER]))
                self.assertEqual([[None], [1], [0]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[sqlite3_column_int]))
                self.assertEqual([[None], [1], [0]], db_rows)

    def test_double(self):
        with sqlite_open(self._path, "crw") as db:
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a DOUBLE)"))
            with db.transaction():
                with db.prepare("INSERT INTO t VALUES(?)") as stmt:
                    for value in (None, 1.0, 0.0):
                        stmt.reset()
                        stmt.clear_bindings()
                        stmt.bind_double(1, value)
                        try:
                            self.assertEqual(1, stmt.execute_update())
                        except Exception as e:
                            print(e)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query())
                self.assertEqual([[None], [1.0], [0.0]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[SQLITE_FLOAT]))
                self.assertEqual([[None], [1.0], [0.0]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[sqlite3_column_double]))
                self.assertEqual([[None], [1.0], [0.0]], db_rows)

    def test_unix_ts(self):
        paris_summer_timezone = dt.timezone(dt.timedelta(hours=2))
        d1 = dt.datetime(2010, 5, 1, 0, 43, 14, tzinfo=paris_summer_timezone)
        d2 = dt.datetime(2003, 6, 3, 14, 1, 21, tzinfo=dt.timezone.utc)

        with sqlite_open(self._path, "crw") as db:
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a DOUBLE)"))
            with db.transaction():
                with db.prepare("INSERT INTO t VALUES(?)") as stmt:
                    for value in (None, d1, d2):
                        stmt.reset()
                        stmt.clear_bindings()
                        stmt.bind_unix_ts(1, value)
                        try:
                            self.assertEqual(1, stmt.execute_update())
                        except Exception as e:
                            print(e)

            with db.prepare("SELECT datetime(a, 'unixepoch') FROM t") as stmt:
                db_rows = list(stmt.execute_query())
                self.assertEqual([[None], ['2010-04-30 22:43:14'], ['2003-06-03 14:01:21']], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query())
                self.assertEqual([[None], [d1.timestamp()], [d2.timestamp()]], db_rows)
                self.assertEqual([[None], [1272667394.0], [1054648881.0]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[PY4LO_UNIX_TS]))
                self.assertEqual([
                    [None],
                    [dt.datetime(2010, 4, 30, 22, 43, 14, tzinfo=dt.timezone.utc)],
                    [dt.datetime(2003, 6, 3, 14, 1, 21, tzinfo=dt.timezone.utc)]
                ], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[decode_unix_ts_to_datetime_utc]))
                self.assertEqual([
                    [None],
                    [dt.datetime(2010, 4, 30, 22, 43, 14, tzinfo=dt.timezone.utc)],
                    [dt.datetime(2003, 6, 3, 14, 1, 21, tzinfo=dt.timezone.utc)]
                ], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[create_decode_unix_ts_to_datetime(paris_summer_timezone)]))
                self.assertEqual([
                    [None],
                    [dt.datetime(2010, 5, 1, 0, 43, 14, tzinfo=paris_summer_timezone)],
                    [dt.datetime(2003, 6, 3, 16, 1, 21, tzinfo=paris_summer_timezone)]
                ], db_rows)

    def test_julian(self):
        paris_summer_timezone = dt.timezone(dt.timedelta(hours=2))
        d1 = dt.datetime(2010, 5, 1, 0, 43, 14, tzinfo=paris_summer_timezone)
        d2 = dt.datetime(2003, 6, 3, 14, 1, 21, tzinfo=dt.timezone.utc)

        with sqlite_open(self._path, "crw") as db:
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a DOUBLE)"))
            with db.transaction():
                with db.prepare("INSERT INTO t VALUES(?)") as stmt:
                    for value in (None, d1, d2):
                        stmt.reset()
                        stmt.clear_bindings()
                        stmt.bind_julian(1, value)
                        try:
                            self.assertEqual(1, stmt.execute_update())
                        except Exception as e:
                            print(e)

            with db.prepare("SELECT datetime(a) FROM t") as stmt:
                db_rows = list(stmt.execute_query())
                self.assertEqual([[None], ['2010-04-30 22:43:14'], ['2003-06-03 14:01:21']], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query())
                self.assertEqual([[None], [datetime_to_julian(d1)], [datetime_to_julian(d2)]], db_rows)
                self.assertEqual([[None], [2455317.446689815], [2452794.0842708335]], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[PY4LO_JULIAN]))
                self.assertEqual([
                    [None],
                    [dt.datetime(2010, 4, 30, 22, 43, 13, 999998, tzinfo=dt.timezone.utc)],
                    [dt.datetime(2003, 6, 3, 14, 1, 21, 17, tzinfo=dt.timezone.utc)]
                ], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[decode_julian_to_datetime_utc]))
                self.assertEqual([
                    [None],
                    [dt.datetime(2010, 4, 30, 22, 43, 13, 999998, tzinfo=dt.timezone.utc)],
                    [dt.datetime(2003, 6, 3, 14, 1, 21, 17, tzinfo=dt.timezone.utc)]
                ], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[create_decode_julian_to_datetime(paris_summer_timezone)]))
                self.assertEqual([
                    [None],
                    [dt.datetime(2010, 5, 1, 0, 43, 13, 999998, tzinfo=paris_summer_timezone)],
                    [dt.datetime(2003, 6, 3, 16, 1, 21, 17, tzinfo=paris_summer_timezone)]
                ], db_rows)

    def test_iso8601(self):
        paris_summer_timezone = dt.timezone(dt.timedelta(hours=2))
        d1 = dt.datetime(2010, 5, 1, 0, 43, 14, tzinfo=paris_summer_timezone)
        d2 = dt.datetime(2003, 6, 3, 14, 1, 21, tzinfo=dt.timezone.utc)

        with sqlite_open(self._path, "crw") as db:
            self.assertEqual(0, db.execute_update(
                "CREATE TABLE t(a TEXT)"))
            with db.transaction():
                with db.prepare("INSERT INTO t VALUES(?)") as stmt:
                    for value in (None, d1, d2):
                        stmt.reset()
                        stmt.clear_bindings()
                        stmt.bind_iso8601(1, value)
                        try:
                            self.assertEqual(1, stmt.execute_update())
                        except Exception as e:
                            print(e)

            with db.prepare("SELECT datetime(a) FROM t") as stmt:
                db_rows = list(stmt.execute_query())
                self.assertEqual([[None], ['2010-04-30 22:43:14'], ['2003-06-03 14:01:21']], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query())
                self.assertEqual([[None], [d1.isoformat()], [d2.isoformat()]], db_rows)
                self.assertEqual([[None], ['2010-05-01T00:43:14+02:00'], ['2003-06-03T14:01:21+00:00']], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[PY4LO_ISO8601]))
                self.assertEqual([
                    [None],
                    [dt.datetime(2010, 4, 30, 22, 43, 14, tzinfo=dt.timezone.utc)],
                    [dt.datetime(2003, 6, 3, 14, 1, 21, tzinfo=dt.timezone.utc)]
                ], db_rows)

            with db.prepare("SELECT * FROM t") as stmt:
                db_rows = list(stmt.execute_query(column_decodes=[decode_iso8601_to_datetime]))
                self.assertEqual([
                    [None],
                    [dt.datetime(2010, 4, 30, 22, 43, 14, tzinfo=dt.timezone.utc)],
                    [dt.datetime(2003, 6, 3, 14, 1, 21, tzinfo=dt.timezone.utc)]
                ], db_rows)
                self.assertEqual([
                    [paris_summer_timezone],
                    [dt.timezone.utc]
                ], [[d.tzinfo for d in row] for row in db_rows[1:]])


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

    def test_example(self):
        with sqlite_open(":memory:", "crw") as db:
            db.execute_update(
                "CREATE TABLE identifiers(identifier INTEGER, name TEXT, date DOUBLE)")

            with db.transaction():
                with db.prepare(
                        "INSERT INTO identifiers VALUES(?, ?, ?)") as stmt:
                    stmt.reset()
                    stmt.clear_bindings()
                    stmt.bind_int(1, 1)
                    stmt.bind_text(2, "foo")
                    stmt.bind_unix_ts(3, dt.datetime(
                        2024, 12, 14, 13, 20, 59, tzinfo=dt.timezone.utc))
                    try:
                        stmt.execute_update()  # returns 1
                    except Exception as e:
                        print(e)

            decodes = [
                SQLITE_INTEGER,
                SQLITE_TEXT,
                create_decode_unix_ts_to_datetime(
                    dt.timezone(dt.timedelta(hours=1)))
                # we want a specific timezone
            ]

            with db.prepare("SELECT * FROM identifiers") as stmt:
                for db_row in stmt.execute_query(
                        with_names=True, column_decodes=decodes):
                    print(db_row)
                    # {'identifier': 1, 'name': 'foo',
                    #  'date': datetime.datetime(
                    #     2024, 12, 14, 14, 20, 59,
                    #     tzinfo=datetime.timezone(datetime.timedelta(seconds=3600))
                    # )}


if __name__ == '__main__':
    unittest.main()
