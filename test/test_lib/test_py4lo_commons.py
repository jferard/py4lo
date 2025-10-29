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
import configparser
import datetime as dt
import os
import subprocess  # nosec: B404
import tempfile
import unittest
import zipfile
from decimal import Decimal
from io import TextIOWrapper
from pathlib import Path
from unittest import mock

import py4lo_commons
from py4lo_commons import (
    secure_strip,
    uno_url_to_path, uno_path_to_url, create_bus, Commons, init, sanitize,
    read_config, create_parse_int_or, create_format_int_or,
    create_parse_float_or, create_parse_decimal_or, create_parse_date_or,
    create_parse_datetime_or, create_parse_time_or, create_format_float_or,
    create_format_decimal_or, create_format_date_or, create_format_datetime_or,
    create_format_time_or
)


class MiscTestCase(unittest.TestCase):
    def test_uno(self):
        from _mock_objects import uno
        self.assertEqual("url/", uno.fileUrlToSystemPath("file://url"))
        self.assertEqual("file:///url", uno.systemPathToFileUrl("/url"))

    @mock.patch("py4lo_commons.uno.fileUrlToSystemPath")
    def test_uno_url_to_path(self, futsp):
        # prepare
        futsp.side_effect = ["path"]

        # play
        ret = uno_url_to_path("abc")

        # verify
        self.assertEqual(Path("path"), ret)

    @mock.patch("py4lo_commons.uno.fileUrlToSystemPath")
    def test_uno_url_to_path_empty(self, futsp):
        # prepare
        futsp.side_effect = ["path"]

        # play
        ret = uno_url_to_path("  ")

        # verify
        self.assertIsNone(ret)

    @mock.patch("py4lo_commons.uno.systemPathToFileUrl")
    def test_uno_path_to_url(self, sptfu):
        # prepare
        sptfu.side_effect = ["url"]

        # play
        ret = uno_path_to_url("abc")

        # verify
        self.assertEqual("url", ret)
        self.assertEqual([mock.call(str(Path("abc").resolve()))],
                         sptfu.mock_calls)

    @mock.patch("py4lo_commons.uno.systemPathToFileUrl")
    def test_uno_path_to_url_err(self, sptfu):
        # prepare
        sptfu.side_effect = ["url"]
        path = mock.Mock()
        path.resolve.side_effect = FileNotFoundError

        # play
        ret = uno_path_to_url(path)

        # verify
        self.assertEqual("url", ret)
        self.assertEqual([mock.call(str(path))], sptfu.mock_calls)


class TestBus(unittest.TestCase):
    def setUp(self):
        self.b = create_bus()
        self.s = mock.Mock()

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
        self.assertEqual(py4lo_commons._xsc, xsc)

    def test_create_linux(self):
        oDoc = mock.Mock(URL="file:///doc/url.ods")

        xsc = mock.Mock()
        xsc.getDocument.side_effect = [oDoc]

        init(xsc)
        c = Commons.create()
        self.assertEqual(Path("/doc"), c.cur_dir())

    def test_create_linux2(self):
        # prepare
        oDoc = mock.Mock(URL="file:///doc/url.ods")

        xsc = mock.Mock()
        xsc.getDocument.side_effect = [oDoc]

        # play
        c = Commons.create(xsc)

        # verify
        self.assertEqual(Path("/doc"), c.cur_dir())

    # def test_create_win(self):
    #     oDoc = mock.Mock(URL="file://C:/doc/url.ods")
    #
    #     xsc = mock.Mock()
    #     xsc.getDocument.side_effect = [oDoc]
    #
    #     init(xsc)
    #     oTextRange = Commons.create()
    #     self.assertEqual(Path("C:/doc"), oTextRange.cur_dir())

    def testCurDir(self):
        self.assertEqual(Path("/temp"), self.c.cur_dir())

    def test_sanitize(self):
        self.assertEqual("e", sanitize("é"))

    def test_secure_strip(self):
        self.assertEqual("x", secure_strip("x"))
        self.assertEqual("x", secure_strip("x "))
        self.assertEqual("x", secure_strip(" x"))
        self.assertEqual("x", secure_strip(" x "))
        self.assertEqual(10, secure_strip(10))

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

    def test_logger_err(self):
        t = tempfile.mkdtemp()
        c = Commons((Path(t) / "file.ods").as_uri())
        with self.assertRaises(Exception):
            _ = c.logger()

        os.rmdir(t)

    def test_read_internal_config(self):
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

        config2 = c.read_internal_config("config", {})
        self.assertEqual(config, config2)
        path.unlink()
        os.rmdir(t)

    def test_read_internal_config_missing(self):
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

        config2 = c.read_internal_config(["config", "config2"], {})
        self.assertEqual(config, config2)
        path.unlink()
        os.rmdir(t)

    def test_read_asset(self):
        t = tempfile.mkdtemp()
        path = Path(t) / "file.ods"
        c = Commons(path.as_uri())
        zf = zipfile.ZipFile(path, "w")
        try:
            s = zf.open("asset", "w")
            try:
                s.write(b"abc")
            finally:
                s.close()
        finally:
            zf.close()

        self.assertEqual(b"abc", c.get_asset("asset"))
        path.unlink()
        os.rmdir(t)

    def test_read_empty_config(self):
        config = read_config([])
        self.assertEqual(['DEFAULT'], list(config))

    def test_get_logger(self):
        t = tempfile.mkdtemp()
        c = Commons((Path(t) / "file.ods").as_uri())
        logger = c.get_logger()
        logger.debug("test")

        with (Path(t) / "py4lo.log").open("r") as s:
            text = s.read()
            self.assertTrue(text.endswith(" - root - DEBUG - test\n"))


class ParseTestCase(unittest.TestCase):
    def test_parse_int1(self):
        parse_int = create_parse_int_or(" ", None)
        self.assertEqual(12345, parse_int("12345"))
        self.assertEqual(12345, parse_int("12 345"))
        self.assertEqual(-12345, parse_int("-12 345"))
        self.assertEqual(12345, parse_int("+12 345"))
        self.assertIsNone(parse_int("12,345"))
        self.assertIsNone(parse_int("foo"))

    def test_parse_int2(self):
        parse_int = create_parse_int_or(" ,", None)
        self.assertEqual(12345, parse_int("12345"))
        self.assertEqual(12345, parse_int("12 345"))
        self.assertEqual(12345, parse_int("12,345"))
        self.assertEqual(-12345, parse_int("-12,345"))
        self.assertEqual(12345, parse_int("+12,345"))
        self.assertEqual(12345789, parse_int("12,345 789"))
        self.assertIsNone(parse_int("foo"))

    def test_parse_float1(self):
        parse_float = create_parse_float_or(",", ".", None)
        self.assertAlmostEqual(12345.678, parse_float("12345.678"))
        self.assertAlmostEqual(12345.678, parse_float("12,345.678"))
        self.assertAlmostEqual(-12345.678, parse_float("-12,345.678"))
        self.assertAlmostEqual(12345.678, parse_float("+12,345.678"))
        self.assertIsNone(parse_float("12 345.678"))
        self.assertIsNone(parse_float("foo"))

    def test_parse_float2(self):
        parse_float = create_parse_float_or("\xa0. ", ",", None)
        self.assertAlmostEqual(12345.678, parse_float("12345,678"))
        self.assertAlmostEqual(12345.678, parse_float("12\xa0345,678"))
        self.assertAlmostEqual(12345.678, parse_float("12.345,678"))
        self.assertAlmostEqual(12345.678, parse_float("12 345,678"))
        self.assertAlmostEqual(-12345.678, parse_float("-12 345,678"))
        self.assertAlmostEqual(12345.678, parse_float("+12 345,678"))
        self.assertAlmostEqual(12345678.0, parse_float("12 345.678"))
        self.assertIsNone(parse_float("foo"))

    def test_parse_decimal1(self):
        parse_decimal = create_parse_decimal_or(",", ".", None)
        self.assertEqual(Decimal("12345.678"), parse_decimal("12345.678"))
        self.assertEqual(Decimal("12345.678"), parse_decimal("12,345.678"))
        self.assertEqual(Decimal("-12345.678"), parse_decimal("-12,345.678"))
        self.assertEqual(Decimal("12345.678"), parse_decimal("+12,345.678"))
        self.assertIsNone(parse_decimal("12 345.678"))
        self.assertIsNone(parse_decimal("foo"))

    def test_parse_decimal2(self):
        parse_decimal = create_parse_decimal_or("\xa0. ", ",", None)
        self.assertEqual(Decimal("12345.678"), parse_decimal("12345,678"))
        self.assertEqual(Decimal("12345.678"), parse_decimal("12\xa0345,678"))
        self.assertEqual(Decimal("12345.678"), parse_decimal("12.345,678"))
        self.assertEqual(Decimal("12345.678"), parse_decimal("12 345,678"))
        self.assertEqual(Decimal("-12345.678"), parse_decimal("-12 345,678"))
        self.assertEqual(Decimal("12345.678"), parse_decimal("+12 345,678"))
        self.assertEqual(Decimal("12345678"), parse_decimal("12 345.678"))
        self.assertIsNone(parse_decimal("foo"))

    def test_parse_date(self):
        parse_date = create_parse_date_or("%d/%m/%Y", "NULL")
        self.assertEqual(dt.date(2025, 10, 25), parse_date("25/10/2025"))
        self.assertEqual("NULL", parse_date("1/25/2025"))

    def test_parse_datetime(self):
        parse_datetime = create_parse_datetime_or("%d/%m/%Y %H:%M:%S", "NULL")
        self.assertEqual(dt.datetime(2025, 10, 25, 17, 40, 53),
                         parse_datetime("25/10/2025 17:40:53"))
        self.assertEqual("NULL", parse_datetime("1/25/2025 17:40:53"))

    def test_parse_time(self):
        parse_time = create_parse_time_or("%H:%M:%S", "NULL")
        self.assertEqual(dt.time(17, 40, 53),
                         parse_time("17:40:53"))
        self.assertEqual("NULL", parse_time("17-40-53"))


class FormatTestCase(unittest.TestCase):
    def test_format_int1(self):
        format_int = create_format_int_or("\xa0", "NULL")
        self.assertEqual('123', format_int(123))
        self.assertEqual('12\xa0345', format_int(12345))
        self.assertEqual("NULL", format_int(None))

    def test_format_int2(self):
        format_int = create_format_int_or(",", "NULL")
        self.assertEqual('123', format_int(123))
        self.assertEqual('12,345', format_int(12345))
        self.assertEqual("NULL", format_int(None))

    def test_format_float(self):
        format_float = create_format_float_or("", ".", -1, "NULL")
        self.assertEqual('12345.678', format_float(12345.678))
        self.assertEqual('NULL', format_float(None))

        format_float = create_format_float_or("", ",", -1, "NULL")
        self.assertEqual('12345,678', format_float(12345.678))
        self.assertEqual('NULL', format_float(None))

        format_float = create_format_float_or("_", ",", 2, "NULL")
        self.assertEqual('12_345,68', format_float(12345.678))
        self.assertEqual('NULL', format_float(None))

        format_float = create_format_float_or("_", ".", -1, "NULL")
        self.assertEqual('12_345.678', format_float(12345.678))
        self.assertEqual('NULL', format_float(None))

        format_float = create_format_float_or(",", ",", -1, "NULL")
        self.assertEqual('12,345,678', format_float(12345.678))
        self.assertEqual('NULL', format_float(None))

        format_float = create_format_float_or("\xa0", ",", 2, "NULL")
        self.assertEqual('12\xa0345,68', format_float(12345.678))
        self.assertEqual('NULL', format_float(None))

        format_float = create_format_float_or("\xa0", ".", -1, "NULL")
        self.assertEqual('12\xa0345.678', format_float(12345.678))
        self.assertEqual('NULL', format_float(None))

        format_float = create_format_float_or("\xa0", ",", -1, "NULL")
        self.assertEqual('12\xa0345,678', format_float(12345.678))
        self.assertEqual('NULL', format_float(None))

        format_float = create_format_float_or("\xa0", ".", 2, "NULL")
        self.assertEqual('12\xa0345.68', format_float(12345.678))

        format_float = create_format_float_or("_", ".", 2, "NULL")
        self.assertEqual('12_345.68', format_float(12345.678))

        format_float = create_format_float_or("", ",", 2, "NULL")
        self.assertEqual('12345,68', format_float(12345.678))

        format_float = create_format_float_or("", ".", 2, "NULL")
        self.assertEqual('12345.68', format_float(12345.678))

    def test_format_decimal(self):
        format_decimal = create_format_decimal_or("_", ",", 2, "NULL")
        self.assertEqual('12_345,68', format_decimal(Decimal("12345.678")))
        self.assertEqual('-12_345,68', format_decimal(Decimal("-12345.678")))
        self.assertEqual('NULL', format_decimal(None))

        format_decimal = create_format_decimal_or("\xa0", ",", 2, "NULL")
        self.assertEqual('12\xa0345,68', format_decimal(Decimal("12345.678")))
        self.assertEqual('-12\xa0345,68', format_decimal(Decimal("-12345.678")))
        self.assertEqual('NULL', format_decimal(None))

        format_decimal = create_format_decimal_or("_", ",", -1, "NULL")
        self.assertEqual('12_345,0', format_decimal(Decimal("12345")))
        self.assertEqual('12_345,0', format_decimal(Decimal("12345.0")))
        self.assertEqual('12_345,678', format_decimal(Decimal("12345.678")))
        self.assertEqual('NULL', format_decimal(None))

    def test_format_date(self):
        format_date = create_format_date_or("%d/%m/%Y", "NULL")
        self.assertEqual('29/10/2025', format_date(dt.date(2025, 10, 29)))
        self.assertEqual("NULL", format_date(None))

    def test_format_datetime(self):
        format_datetime = create_format_datetime_or("%d/%m/%Y %H:%M:%S", "NULL")
        self.assertEqual('29/10/2025 19:50:04', format_datetime(dt.datetime(2025, 10, 29, 19, 50, 4)))
        self.assertEqual("NULL", format_datetime(None))

    def test_format_time(self):
        format_time = create_format_time_or("%H:%M:%S", "NULL")
        self.assertEqual('19:50:04', format_time(dt.time(19, 50, 4)))
        self.assertEqual("NULL", format_time(None))


if __name__ == '__main__':
    unittest.main()
