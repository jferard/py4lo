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
import subprocess   # nosec: B404
import tempfile
import unittest
import zipfile
from io import TextIOWrapper
import os
from pathlib import Path

from unittest import mock

import py4lo_commons
from py4lo_commons import (
    uno_url_to_path, uno_path_to_url, create_bus, Commons, init, sanitize,
    read_config)


class MiscTestCase(unittest.TestCase):
    def test_uno(self):
        from _mock_constants import uno
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


if __name__ == '__main__':
    unittest.main()
