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
from unittest.mock import Mock, patch, call, MagicMock

from script_set_processor import *


class TestScriptsProcessor(unittest.TestCase):
    def test(self):
        logger = Mock()
        dp = Mock()
        path = Mock(stem='fname')
        target_dir = Mock()
        bound = MagicMock()
        path.open.return_value = bound
        bound.__enter__.return_value = ["some line"]

        sp = ScriptProcessor(logger, dp, path, target_dir)
        sp.parse_script()

        self.assertEqual(
            [call.log(10, 'Parsing script: %s (%s)', 'fname', path)],
            logger.mock_calls)
        self.assertEqual([call.ignore_lines(), call.end()], dp.mock_calls)
        self.assertEqual(
            [call.open('r', encoding='utf-8'), call.open().__enter__(),
             call.open().__exit__(None, None, None)], path.mock_calls)
