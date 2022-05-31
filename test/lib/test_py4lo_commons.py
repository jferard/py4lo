# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   FastODS is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   FastODS is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""
import unittest
import zipfile
from io import TextIOWrapper
from unittest import mock
from unittest.mock import *
import tempfile
import os
import subprocess
import tst_env
import sys

from py4lo_commons import *
from py4lo_commons import _get_config


class TestBus(unittest.TestCase):
    def setUp(self):
        self.b = Bus()
        self.s = Mock()

    def test(self):
        self.b.subscribe(self.s)
        self.b.post("b", "a")
        self.b.post("a", "b")

        # verify
        self.s._handle_b.assert_called_with("a")
        self.s._handle_a.assert_called_with("b")

    def test2(self):
        self.b.subscribe(self)
        self.b.post("x", "a")
        self.b.post("y", "a")
        self.assertEqual("oka", self.v)

    def _handle_y(self, data):
        self.v = "ok" + data


class TestCommons(unittest.TestCase):
    def setUp(self):
        self.c = Commons("file:///temp/file.ods")

    def test_init(self):
        xsc = object()
        init(xsc)
        self.assertEqual(Commons.xsc, xsc)

    def test_create_linux(self):
        oDoc = mock.Mock(URL="file:///doc/url.ods")

        xsc = mock.Mock()
        xsc.getDocument.side_effect = [oDoc]

        init(xsc)
        c = Commons.create()
        self.assertEqual(Path("/doc"), c.cur_dir())

    # def test_create_win(self):
    #     oDoc = mock.Mock(URL="file://C:/doc/url.ods")
    #
    #     xsc = mock.Mock()
    #     xsc.getDocument.side_effect = [oDoc]
    #
    #     init(xsc)
    #     c = Commons.create()
    #     self.assertEqual(Path("C:/doc"), c.cur_dir())

    def testCurDir(self):
        self.assertEqual(Path("/temp"), self.c.cur_dir())

    def testSanitize(self):
        self.assertEqual("e", sanitize("é"))

    def test_logger(self):
        t = tempfile.NamedTemporaryFile(delete=False, mode='w')
        self.c.init_logger(t.name, fmt='%(levelname)s - %(message)s')
        self.c.logger().debug("é")
        t.flush()
        t.close()
        res = subprocess.getoutput("cat {}".format(t.name))
        self.assertEqual("DEBUG - é", res)
        os.unlink(t.name)

    def test_logger_init_twice(self):
        t = tempfile.NamedTemporaryFile(delete=False, mode='w')
        self.c.init_logger(t, fmt='%(levelname)s - %(message)s')
        with self.assertRaises(Exception):
            self.c.init_logger(t, fmt='%(levelname)s - %(message)s')
        os.unlink(t.name)

    def test_logger_none(self):
        t = tempfile.mkdtemp()
        c = Commons((Path(t) / "file.ods").as_uri())
        logger = c.get_logger()
        logger.debug("test")

        with (Path(t) / "py4lo.log").open("r") as s:
            text = s.read()
            self.assertTrue(text.endswith(" - root - DEBUG - test\n"))

    def test_get_config(self):
        t = tempfile.mkdtemp()
        path = Path(t) / "file.ods"
        c = Commons(path.as_uri())
        config = configparser.ConfigParser()
        config["S"] = {"a": "1"}
        zf = zipfile.ZipFile(path, "w")
        try:
            s = zf.open("config", "w")
            config.write(TextIOWrapper(s, encoding="utf-8"))
        finally:
            zf.close()

        config2 = c.read_internal_config(["config"], {})
        self.assertEqual(config, config2)
        path.unlink()
        os.rmdir(t)

    def test_get_logger(self):
        t = tempfile.mkdtemp()
        c = Commons((Path(t) / "file.ods").as_uri())
        logger = c.get_logger()
        logger.debug("test")

        with (Path(t) / "py4lo.log").open("r") as s:
            text = s.read()
            self.assertTrue(text.endswith(" - root - DEBUG - test\n"))


class TestDate(unittest.TestCase):
    def test_date_to_int(self):
        self.assertEqual(0, date_to_int(dt.datetime(1899, 12, 30)))
        self.assertEqual(44639, date_to_int(dt.datetime(2022, 3, 19)))
        self.assertEqual(44639, date_to_int(dt.date(2022, 3, 19)))

    def test_date_to_float(self):
        self.assertEqual(0.0, date_to_float(dt.datetime(1899, 12, 30)))
        self.assertEqual(0.0, date_to_float(dt.date(1899, 12, 30)))
        self.assertAlmostEqual(44639.723032407404, date_to_float(dt.datetime(
            2022, 3, 19, 17, 21, 10)), delta=0.0001)
        self.assertAlmostEqual(0.723032407404, date_to_float(dt.time(
            17, 21, 10)), delta=0.0001)

    def test_int_to_date(self):
        self.assertEqual(dt.datetime(1899, 12, 30), int_to_date(0))
        self.assertEqual(dt.datetime(2022, 3, 19), int_to_date(44639))
        self.assertEqual(dt.datetime(2022, 3, 19), int_to_date(44639))

    def test_float_to_date(self):
        self.assertEqual(dt.datetime(1899, 12, 30), float_to_date(0.0))
        self.assertEqual(dt.datetime(1899, 12, 30), float_to_date(0.0))
        self.assertEqual(dt.datetime(
            2022, 3, 19, 17, 21, 10), float_to_date(44639.723032407404))
        self.assertEqual(dt.datetime(
            1899, 12, 30, 17, 21, 10), float_to_date(0.723032407404))


if __name__ == '__main__':
    unittest.main()
