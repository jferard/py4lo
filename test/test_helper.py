# -*- coding: utf-8 -*-
# Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2021 J. Férard <https://github.com/jferard>
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

# see http://stackoverflow.com/questions/61151/where-do-the-python-unit-tests-go
import sys
import unittest
from pathlib import Path
from typing import IO
from unittest.mock import Mock, MagicMock, call


def any_object():
    class AnyObject:
        def __eq__(self, other):
            return True

    return AnyObject()


# append module root directory to sys.path
test_dir = Path(__file__).parent
root_dir = test_dir.parent
py4lo_dir = root_dir.joinpath("py4lo")
lib_dir = root_dir.joinpath("lib")
inc_dir = root_dir.joinpath("inc")

for p in map(str, [py4lo_dir, lib_dir, inc_dir]):
    if p not in sys.path:
        sys.path.insert(0, p)


def file_path_mock(content: IO, **kwargs) -> Path:
    s = MagicMock(**kwargs)
    manager = MagicMock()
    s.open.return_value = manager
    manager.__enter__.return_value = content
    return s


def file_path_error_mock(**kwargs):
    def mock_enter(*_args, **_kwargs):
        raise FileNotFoundError("")

    s = MagicMock(**kwargs)
    manager = MagicMock()
    s.open.return_value = manager
    manager.__enter__ = mock_enter
    return s


def verify_open_path(tc: unittest.TestCase, s: Path, *args, **kwargs):
    tc.assertTrue(
        all(x in s.mock_calls for x in [call.open(*args, **kwargs), call.open().__enter__(),
         call.open().__exit__(None, None, None)]))
