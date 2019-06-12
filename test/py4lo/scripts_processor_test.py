# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2019 J. Férard <https://github.com/jferard>

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
from scripts_processor import *
import scripts_processor

class TestScriptsProcessor(unittest.TestCase):
    def setUp(self):
        self.maxDiff = None

    def test_target_script(self):
        e = Exception("d")
        t = TargetScript("a/z", "b", ["c"], e)
        self.assertEqual("a/z", t.get_fname())
        self.assertEqual("z", t.get_name())
        self.assertEqual("b", t.get_content())
        self.assertEqual(["c"], t.get_exported_func_names())
        self.assertEqual(e, t.get_exception())

    @patch("py_compile.compile")
    @patch("builtins.open")
    def test_scripts_processor(self, mock_open, mock_compile):
        logger = Mock()

        src = MagicMock()
        target = MagicMock()
        dest = MagicMock()
        mock_open.side_effect = [src, target]

        src.__enter__.return_value = ["some line"]
        target.__enter__.return_value = dest

        try:
            sp = scripts_processor.ScriptsProcessor(logger, "src", "target", "3.6")
            sp.process(["script_fname"])
        except Exception as e:
            pass

        self.assertEqual([
            call('script_fname', 'r', encoding='utf-8'),
            call('target/script_fname', 'wb'),
        ], mock_open.mock_calls)
        self.assertEqual([call('script_fname', doraise=True)], mock_compile.mock_calls)
        self.assertEqual([call.write(b'# parsed by py4lo (https://github.com/jferard/py4lo)\nsome line')],
            dest.mock_calls)
        self.assertEqual([
            call.log(10, 'Scripts to process: %s', ['script_fname']),
            call.log(10, 'Parsing script: %s (%s)', 'script_fname', 'script_fname'),
            call.log(10, 'Writing script: %s (%s)', 'script_fname', 'target/script_fname'),
        ], logger.mock_calls)


