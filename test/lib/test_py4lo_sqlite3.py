import random
import string
import time
import unittest
from datetime import datetime
from pathlib import Path
from pprint import pprint

from py4lo_sqlite3 import sqlite_open, SQLType


class Sqlite3TestCase(unittest.TestCase):
    def test_sqlite3(self):
        print("-> remove db")
        path = Path("test.sqlite3")
        path.unlink(True)

        with sqlite_open(path, "crw") as db:
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


if __name__ == '__main__':
    unittest.main()
