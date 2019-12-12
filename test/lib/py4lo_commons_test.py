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
from unittest.mock import *
import env
import sys, os

print("commons", sys.path)
from py4lo_commons import *

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
        self.c = Commons("file:///temp/")

    def testCurDir(self):
        self.assertTrue(self.c.cur_dir().endswith("temp"))

    def testSanitize(self):
        self.assertEqual("e", sanitize("é"))

    def testLogger(self):
        import tempfile, os, subprocess
        t = tempfile.NamedTemporaryFile(delete=False, mode='w')
        self.c.init_logger(t, format='%(levelname)s - %(message)s')
        print(self.c.logger().handlers)
        self.c.logger().debug("é")
        t.flush()
        t.close()
        res = subprocess.getoutput("cat {}".format(t.name))
        self.assertEqual("DEBUG - é", res)
        os.unlink(t.name)


class TestOther(unittest.TestCase):
    pass

if __name__ == '__main__':
    unittest.main()
