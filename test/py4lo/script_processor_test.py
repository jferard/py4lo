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

from scripts_processor import *


class TestScriptsProcessor(unittest.TestCase):
    @patch("builtins.open")
    def test(self, mock_open):
        logger = Mock()
        dp = Mock()

        bound = MagicMock()
        mock_open.return_value = bound
        bound.__enter__.return_value = ["some line"]

        sp = ScriptProcessor(logger, dp, "fname")
        sp.parse_script()

        self.assertEqual([call.log(10, 'Parsing script: %s (%s)', 'fname', 'fname')], logger.mock_calls)
        self.assertEqual([call.ignore_lines(), call.end()], dp.mock_calls)
        self.assertEqual([call('fname', 'r', encoding='utf-8'), call().__enter__(), call().__exit__(None, None, None)],
                         mock_open.mock_calls)
        self.assertEqual([call.__enter__(), call.__exit__(None, None, None)],
                         bound.mock_calls)