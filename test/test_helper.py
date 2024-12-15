# -*- coding: utf-8 -*-
# Py4LO - Python Toolkit For LibreOffice Calc
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

# see http://stackoverflow.com/questions/61151/where-do-the-python-unit-tests-go
import sys
import unittest
from pathlib import Path
from typing import IO
from unittest import mock


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
    s = mock.MagicMock(**kwargs)
    manager = mock.MagicMock()
    s.open.return_value = manager
    manager.__enter__.return_value = content
    return s


def file_path_error_mock(**kwargs):
    def mock_enter(*_args, **_kwargs):
        raise FileNotFoundError("")

    s = mock.MagicMock(**kwargs)
    manager = mock.MagicMock()
    s.open.return_value = manager
    manager.__enter__ = mock_enter
    return s


def verify_open_path(tc: unittest.TestCase, s: Path, *args, **kwargs):
    tc.assertTrue(
        all(x in s.mock_calls for x in
            [mock.call.open(*args, **kwargs), mock.call.open().__enter__(),
             mock.call.open().__exit__(None, None, None)]))


def compare_xml_strings(s1, s2):
    """
    >>> compare_xml_strings('<root a="1" b="2">A</root>', '<root b="2" a="1">A</root>')

    @param s1:
    @param s2:
    @return:
    """
    import xml.etree.ElementTree as ET
    dom1 = ET.fromstring(s1)
    dom2 = ET.fromstring(s2)
    stack = [(dom1, dom2)]
    while stack:
        dom1, dom2 = stack.pop()
        if dom1.tag != dom2.tag:
            return False
        if dom1.attrib != dom2.attrib:
            return False
        if dom1.text != dom2.text:
            return False
        c1s = list(dom1)
        c2s = list(dom2)
        if len(c1s) != len(c2s):
            return False
        for c1, c2 in zip(c1s, c2s):
            stack.append((c1, c2))

    return True


if __name__ == "__main__":
    import doctest

    doctest.testmod()
