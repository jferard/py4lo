# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2020 J. Férard <https://github.com/jferard>

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
import tst_env
from branch_processor import BranchProcessor


class TestBranchProcessor(unittest.TestCase):
    def setUp(self):
        self.bp = BranchProcessor(lambda args:args[0])

    def test_false_branch(self):
        self.assertFalse(self.bp.handle_directive("foo", [True]))

    def test_branch(self):
        # before everything
        self.assertFalse(self.bp.skip())
        self.assertTrue(self.bp.handle_directive("if", [True]))
        # in first if
        self.assertFalse(self.bp.skip())
        self.assertTrue(self.bp.handle_directive("if", [False]))
        # in first if, but not in second if
        self.assertTrue(self.bp.skip())
        self.assertTrue(self.bp.handle_directive("elif", [True]))
        # in first if, in second elif
        self.assertFalse(self.bp.skip())
        self.assertTrue(self.bp.handle_directive("endif", [])) # not tested
        # in first if
        self.assertFalse(self.bp.skip())
        self.assertTrue(self.bp.handle_directive("elif", [])) # not tested
        # out of first if : even if condition is true, "el" means "else"
        self.assertTrue(self.bp.skip())
        self.assertTrue(self.bp.handle_directive("endif", [])) # not tested
        # after everything
        self.assertFalse(self.bp.skip())
        self.bp.end()

    def test_not_closed(self):
        self.assertTrue(self.bp.handle_directive("if", [True]))
        with self.assertRaises(ValueError):
            self.bp.end()

    def test_closed(self):
        self.assertTrue(self.bp.handle_directive("if", [True]))
        self.assertTrue(self.bp.handle_directive("else", []))
        self.assertTrue(self.bp.handle_directive("endif", []))
        self.bp.end()

if __name__ == '__main__':
    unittest.main()
