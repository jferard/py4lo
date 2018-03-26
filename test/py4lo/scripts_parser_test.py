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
from scripts_processor import *
import scripts_processor

class TestScriptsParser(unittest.TestCase):
    def setUp(self):
        self.maxDiff = None

    @patch("builtins.open")
    def test_script_parser_normal_line_dont_ignore(self, mock_open):
        logger = Mock()
        dp = Mock()
        dp.ignore_lines.return_value = False
        bound = MagicMock()
        mock_open.return_value = bound
        bound.__enter__.return_value = ["some line"]

        sp = scripts_processor._ScriptParser(logger, dp, "script_fname")
        self.assertEqual(('# parsed by py4lo (https://github.com/jferard/py4lo)\nsome line', []), sp.parse())

        self.assertEqual([], logger.mock_calls)
        self.assertEqual([call.ignore_lines(), call.end()], dp.mock_calls)
        self.assertEqual([call.__enter__(), call.__exit__(None, None, None)], bound.mock_calls)
        self.assertEqual([call('script_fname', 'r', encoding='utf-8'), call().__enter__(), call().__exit__(None, None, None)], mock_open.mock_calls)

    @patch("builtins.open")
    def test_script_parser_normal_line_ignore(self, mock_open):
        logger = Mock()
        dp = Mock()
        dp.ignore_lines.return_value = True
        bound = MagicMock()
        mock_open.return_value = bound
        bound.__enter__.return_value = ["some line"]

        sp = scripts_processor._ScriptParser(logger, dp, "script_fname")
        self.assertEqual(('# parsed by py4lo (https://github.com/jferard/py4lo)\n### py4lo ignore: some line', []), sp.parse())

        self.assertEqual([], logger.mock_calls)
        self.assertEqual([call.ignore_lines(), call.end()], dp.mock_calls)
        self.assertEqual([call.__enter__(), call.__exit__(None, None, None)], bound.mock_calls)
        self.assertEqual([call('script_fname', 'r', encoding='utf-8'), call().__enter__(), call().__exit__(None, None, None)], mock_open.mock_calls)

    @patch("builtins.open")
    def test_script_parser_directve_line(self, mock_open):
        logger = Mock()
        dp = Mock()
        dp.process_line.return_value = ["line1", "line2"]
        bound = MagicMock()
        mock_open.return_value = bound
        bound.__enter__.return_value = ["#some line"]

        sp = scripts_processor._ScriptParser(logger, dp, "script_fname")
        self.assertEqual(('# parsed by py4lo (https://github.com/jferard/py4lo)\nline1\nline2', []), sp.parse())

        self.assertEqual([], logger.mock_calls)
        self.assertEqual([call.process_line('#some line'), call.end()], dp.mock_calls)
        self.assertEqual([call.__enter__(), call.__exit__(None, None, None)], bound.mock_calls)
        self.assertEqual([call('script_fname', 'r', encoding='utf-8'), call().__enter__(), call().__exit__(None, None, None)], mock_open.mock_calls)
