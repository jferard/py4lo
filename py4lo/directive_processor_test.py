# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   FastODS is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   FastODS is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""
import unittest
from directive_processor import is_true, BranchProcessor

class TestIsTrue(unittest.TestCase):

    def test_is_true(self):
        self.assertEqual(False, is_true("2.0", "<", "2.0"))
        self.assertEqual(True, is_true("2.0", "<=", "2.0"))
        self.assertEqual(True, is_true("2.0", "==", "2.0"))
        self.assertEqual(True, is_true("2.0", ">=", "2.0"))
        self.assertEqual(False, is_true("2.0", ">", "2.0"))

        self.assertEqual(True, is_true("2.0", "<", "2.1"))
        self.assertEqual(True, is_true("2.0", "<=", "2.1"))
        self.assertEqual(False, is_true("2.0", "==", "2.1"))
        self.assertEqual(False, is_true("2.0", ">=", "2.1"))
        self.assertEqual(False, is_true("2.0", ">", "2.1"))

        self.assertEqual(False, is_true("2.1", "<", "2.0"))
        self.assertEqual(False, is_true("2.1", "<=", "2.0"))
        self.assertEqual(False, is_true("2.1", "==", "2.0"))
        self.assertEqual(True, is_true("2.1", ">=", "2.0"))
        self.assertEqual(True, is_true("2.1", ">", "2.0"))

    def test_false_branch(self):
        branch_processor = BranchProcessor(lambda args:args[0])
        self.assertFalse(branch_processor.handle_directive("foo", [True]))
        
    def test_branch(self):
        branch_processor = BranchProcessor(lambda args:args[0])
        # before everything
        self.assertFalse(branch_processor.skip())
        self.assertTrue(branch_processor.handle_directive("if", [True]))
        # in first if
        self.assertFalse(branch_processor.skip())
        self.assertTrue(branch_processor.handle_directive("if", [False]))
        # in first if, but not in second if
        self.assertTrue(branch_processor.skip())
        self.assertTrue(branch_processor.handle_directive("elif", [True]))
        # in first if, in second elif
        self.assertFalse(branch_processor.skip())
        self.assertTrue(branch_processor.handle_directive("endif", [])) # not tested
        # in first if
        self.assertFalse(branch_processor.skip())
        self.assertTrue(branch_processor.handle_directive("elif", [])) # not tested
        # out of first if : even if condition is true, "el" means "else"
        self.assertTrue(branch_processor.skip())
        self.assertTrue(branch_processor.handle_directive("endif", [])) # not tested
        # after everything
        self.assertFalse(branch_processor.skip())
        
if __name__ == '__main__':
    unittest.main()