# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2018 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   Py4LO is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   Py4LO is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""

import unittest
from unittest.mock import *
import env

from commands.test_command import TestCommand
import os
import subprocess

class TestCommandTest(unittest.TestCase):
    @patch('subprocess.getstatusoutput', spec=subprocess.getstatusoutput)
    @patch('os.walk', spec=os.walk)
    def test(self, os_walk_mock, subprocess_gso_mock):
        os_walk_mock.return_value = [ ("/a", ("b",), ("c_test.py",)), ("/a/b", (), ("d_test.py",)) ]
        subprocess_gso_mock.side_effect = [(0, "ok"), (1, "not ok")]
        logger = MagicMock()
        tc = TestCommand(logger, "py", "a_dir")
        status = tc.execute()

        logger.assert_has_calls([
            call.info('execute: "py" /a/c_test.py'),
            call.info('output: ok'),
            call.info('execute: "py" /a/b/d_test.py'),
            call.info('output: not ok')
        ])
        subprocess_gso_mock.assert_has_calls([call('"py" /a/c_test.py'), call('"py" /a/b/d_test.py')])
        self.assertEqual((1,), status)

if __name__ == '__main__':
    unittest.main()
