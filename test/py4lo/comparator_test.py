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
import unittest
from comparator import Comparator


class TestComparator(unittest.TestCase):
    def test_simple_check(self):
        c = Comparator()

        self.assertFalse(c.check("2.0", "<", "2.0"))
        self.assertTrue(c.check("2.0", "<=", "2.0"))
        self.assertTrue(c.check("2.0", "==", "2.0"))
        self.assertTrue(c.check("2.0", ">=", "2.0"))
        self.assertFalse(c.check("2.0", ">", "2.0"))

        self.assertTrue(c.check("2.0", "<", "2.1"))
        self.assertTrue(c.check("2.0", "<=", "2.1"))
        self.assertFalse(c.check("2.0", "==", "2.1"))
        self.assertFalse(c.check("2.0", ">=", "2.1"))
        self.assertFalse(c.check("2.0", ">", "2.1"))

        self.assertFalse(c.check("2.1", "<", "2.0"))
        self.assertFalse(c.check("2.1", "<=", "2.0"))
        self.assertFalse(c.check("2.1", "==", "2.0"))
        self.assertTrue(c.check("2.1", ">=", "2.0"))
        self.assertTrue(c.check("2.1", ">", "2.0"))

    def test_check_non_existing_var(self):
        c = Comparator()

        with self.assertRaises(NameError):
            c.check("$x", "<", "2.0")

    def test_check_existing_var_but_number(self):
        c = Comparator({"x": 1})

        with self.assertRaises(TypeError):
            self.assertTrue(c.check("$x", "<", "2.0"))

    def test_check_numbers(self):
        c = Comparator()
        self.assertTrue(c.check("2.1f", ">", "2.0i"))


if __name__ == "__main__":
    unittest.main()
