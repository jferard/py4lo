# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
#
#     This file is part of Py4LO.
#
#     Py4LO is free software: you can redistribute it and/or modify
#     it under the terms of the GNU General Public License as published by
#     the Free Software Foundation, either version 3 of the License, or
#     (at your option) any later version.
#
#     Py4LO is distributed in the hope that it will be useful,
#     but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#     GNU General Public License for more details.
#
#     You should have received a copy of the GNU General Public License
#     along with this program.  If not, see <http://www.gnu.org/licenses/>.
import io
import unittest
from unittest import mock

import script_set_processor
from core.script import ParsedScriptContent
from test.test_helper import file_path_mock, verify_open_path


class TestScriptParser(unittest.TestCase):
    def setUp(self):
        self.maxDiff = None
        self._logger = mock.Mock()
        self._dp = mock.Mock()

    def test_script_parser_normal_line_dont_ignore(self):
        self._dp.ignore_lines.return_value = False
        path = file_path_mock(io.StringIO("some line"))

        sp = script_set_processor._ContentParser(self._logger, self._dp, path)
        self.assertEqual(
            ParsedScriptContent(
                """# parsed by py4lo (https://github.com/jferard/py4lo)
some line""",
                []), sp.parse(True))

        self.assertEqual([], self._logger.mock_calls)
        self.assertEqual([mock.call.ignore_lines(), mock.call.end()],
                         self._dp.mock_calls)
        self.verify_open(path)

    def verify_open(self, path):
        verify_open_path(self, path, 'r', encoding="utf-8")

    def test_script_parser_normal_line_ignore(self):
        self._dp.ignore_lines.return_value = True
        path = file_path_mock(io.StringIO("some line"))

        sp = script_set_processor._ContentParser(self._logger, self._dp, path)
        line = '# parsed by py4lo (https://github.com/jferard/py4lo)\n### py4lo ignore: some line'  # noqa: E501
        self.assertEqual(ParsedScriptContent(
            line, []), sp.parse(True))

        self.assertEqual([], self._logger.mock_calls)
        self.assertEqual([mock.call.ignore_lines(), mock.call.end()],
                         self._dp.mock_calls)
        self.verify_open(path)

    def test_script_parser_directve_line(self):
        self._dp.process_line.return_value = ["line1", "line2"]
        path = file_path_mock(io.StringIO("#some line"))

        sp = script_set_processor._ContentParser(self._logger, self._dp, path)
        self.assertEqual(ParsedScriptContent(
            '# parsed by py4lo (https://github.com/jferard/py4lo)\nline1\nline2',
            []), sp.parse(True))

        self.assertEqual([], self._logger.mock_calls)
        self.assertEqual([mock.call.process_line('#some line'),
                          mock.call.end()],
                         self._dp.mock_calls)
        self.verify_open(path)

    def test_one_line_function(self):
        self._dp.ignore_lines.return_value = False
        path = file_path_mock(io.StringIO("def f(x): return x"))

        sp = script_set_processor._ContentParser(self._logger, self._dp, path)
        self.assertEqual(ParsedScriptContent("""# parsed by py4lo (https://github.com/jferard/py4lo)
def f(x): return x


g_exportedScripts = (f,)""", ["f"]), sp.parse(True))
        self.verify_open(path)

    def test_one_public_function(self):
        self._dp.ignore_lines.return_value = False
        path = file_path_mock(io.StringIO("def f(x):\n    return x"))

        sp = script_set_processor._ContentParser(self._logger, self._dp, path)
        self.assertEqual(ParsedScriptContent("""# parsed by py4lo (https://github.com/jferard/py4lo)
def f(x):
    return x


g_exportedScripts = (f,)""", ["f"]), sp.parse(True))
        self.verify_open(path)
