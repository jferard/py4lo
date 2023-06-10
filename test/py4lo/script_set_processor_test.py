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
from logging import Logger
from unittest.mock import *
import tst_env
from script_set_processor import *
import script_set_processor

from tst_env import file_path_mock, verify_open_path


class TestScriptSetProcessor(unittest.TestCase):
    def setUp(self):
        self.maxDiff = None

    @patch("py_compile.compile")
    def test_scripts_processor(self, mock_compile):
        logger: Logger = Mock()
        source_path = file_path_mock(io.StringIO("some line"),
                                     stem='script_fname')

        dest = io.BytesIO()
        target_path = file_path_mock(dest)
        target_dir: Path = Mock()
        target_dir.joinpath.return_value = target_path
        script: SourceScript = MagicMock(relative_path=Path("rel source"),
                                         script_path=source_path)
        dp: DirectiveProvider = Mock()
        target_path.relative_to.side_effect = [Path("rel target")]

        try:
            sp = script_set_processor.ScriptSetProcessor(logger, target_dir,
                                                         "3.7", dp, [script])
            sp.process()
        except Exception as e:
            raise

        self.assertEqual([
            call.debug('Scripts to process: %s', [script]),
            call.debug('Parsing script: %s (%s)', Path("rel source"),
                       script),
            call.debug('Temp output script is: %s (%s)', target_path, []),
            call.debug('Writing temp script: %s (%s)', Path('rel target'),
                       target_path),
        ], logger.mock_calls)
        self.assertEqual([call(str(source_path), doraise=True)],
                         mock_compile.mock_calls)
        self.assertEqual(
            b'# parsed by py4lo (https://github.com/jferard/py4lo)\nsome line',
            dest.getbuffer())
        verify_open_path(self, source_path, 'r', encoding='utf-8')
