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
import subprocess
import unittest
from pathlib import Path
from unittest import mock

from tools import open_with_calc, nested_merge


class TestTools(unittest.TestCase):
    @mock.patch("subprocess.call", spec=subprocess.call)
    def test_open_with_calc(self, subprocess_call_mock):
        open_with_calc(Path("myfile.ods"), "mycalc.exe")
        self.assertEqual(
            [mock.call(["mycalc.exe", "myfile.ods"])], subprocess_call_mock.mock_calls
        )

    def test_merge(self):
        for expected, d1, d2, func in [
            ({"a": 4}, {"a": 1}, {"a": 2}, lambda x: x * 2),
            ({"a": 1, "b": 4}, {"a": 1}, {"b": 2}, lambda x: x * 2),
            ({"a": 1, "b": [4, 6, 8]}, {"a": 1}, {"b": [2, 3, 4]}, lambda x: x * 2),
            (
                {"a": {"b": [2, 4, 6, 8]}},
                {"a": {"b": [1]}},
                {"a": {"b": [2, 3, 4]}},
                lambda x: x * 2,
            ),
        ]:
            self.assertEqual(expected, nested_merge(d1, d2, func))
