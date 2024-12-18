# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2024 J. Férard <https://github.com/jferard>
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
from logging import Logger
from pathlib import Path
from unittest import mock

from core.script import SourceScript
from directive_processor import DirectiveProcessor
from script_set_processor import ScriptProcessor
from test.test_helper import file_path_mock, verify_open_path


class TestScriptsProcessor(unittest.TestCase):
    def test(self):
        logger: Logger = mock.Mock()
        dp: DirectiveProcessor = mock.Mock()
        source_script: SourceScript = mock.Mock(
            relative_path="fname",
            script_path=file_path_mock(
                io.StringIO("some line")),
            export_funcs=True)
        target_dir: Path = mock.Mock()
        target_dir.joinpath.side_effect = [Path('t')]

        sp = ScriptProcessor(logger, dp, source_script, target_dir)
        sp.parse_script()

        self.assertEqual(
            [mock.call.debug('Parsing script: %s (%s)', 'fname',
                        source_script),
             mock.call.debug('Temp output script is: %s (%s)', Path('t'), [])],
            logger.mock_calls)
        self.assertEqual([mock.call.ignore_lines(), mock.call.end()], dp.mock_calls)
        verify_open_path(self, source_script.script_path, 'r', encoding='utf-8')
