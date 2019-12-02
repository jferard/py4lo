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
from script_set_processor import *
import script_set_processor


class TestScriptSetProcessor(unittest.TestCase):
    def setUp(self):
        self.maxDiff = None

    @patch("py_compile.compile")
    def test_scripts_processor(self, mock_compile):
        logger = Mock()
        source_path = Mock(stem='script_fname')
        src = MagicMock()
        target = MagicMock()
        source_path.open.return_value = src

        target_path = Mock(stem='target/script_fname')
        target_dir = Mock()
        target_dir.joinpath.return_value = target_path
        #        target_dir.join_path().stem = 'ok'
        target_path.open.return_value = target
        dest = MagicMock()

        src.__enter__.return_value = ["some line"]
        target.__enter__.return_value = dest

        try:
            sp = script_set_processor.ScriptSetProcessor(logger, Path("src"),
                                                         target_dir,
                                                         "3.7",
                                                         [source_path])
            sp.process()
        except Exception as e:
            pass

        self.assertEqual([
            call.log(10, 'Scripts to process: %s', [source_path]),
            call.log(10, 'Parsing script: %s (%s)', 'script_fname',
                     source_path),
            call.log(10, 'Writing script: %s (%s)', 'target/script_fname',
                     target_path),
        ], logger.mock_calls)
        self.assertEqual([
            call.open('r', encoding='utf-8'),
            call.open().__enter__(),
            call.open().__exit__(None, None, None),
        ], source_path.mock_calls)
        self.assertEqual([call(source_path, doraise=True)],
                         mock_compile.mock_calls)
        self.assertEqual(
            [call.write(
                b'# parsed by py4lo (https://github.com/jferard/py4lo)\nsome line')],
            dest.mock_calls)
