# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2018 J. Férard <https://github.com/jferard>

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
from directive_processor import *

def Any():
    class Any():
        def __eq__(self, other):
            return True
    return Any()

class TestDirectiveProcessor(unittest.TestCase):
    def setUp(self):
        self.__scripts_path = Mock()
        self.__scripts_processor = Mock()
        self.__branch_processor = Mock()
        self.__directive_provider = Mock()
        self.__dp = DirectiveProcessor(self.__scripts_path, self.__scripts_processor, self.__branch_processor, self.__directive_provider)


    def test_append_script(self):
        self.__dp.append_script("ok")

        self.assertEqual([], self.__scripts_path.mock_calls)
        self.assertEqual([call.append_script("ok")], self.__scripts_processor.mock_calls)
        self.assertEqual([], self.__branch_processor.mock_calls)
        self.assertEqual([], self.__directive_provider.mock_calls)

    def test_process_line_comment(self):
        self.__branch_processor.skip.return_value = True

        l = self.__dp.process_line("ok")
        self.assertEqual(["### ok"], l)

        self.assertEqual([], self.__scripts_path.mock_calls)
        self.assertEqual([], self.__scripts_processor.mock_calls)
        self.assertEqual([call.skip()], self.__branch_processor.mock_calls)
        self.assertEqual([], self.__directive_provider.mock_calls)

    def test_process_line_write(self):
        self.__branch_processor.skip.return_value = False

        l = self.__dp.process_line("ok")
        self.assertEqual(["ok"], l)

        self.assertEqual([], self.__scripts_path.mock_calls)
        self.assertEqual([], self.__scripts_processor.mock_calls)
        self.assertEqual([call.skip()], self.__branch_processor.mock_calls)
        self.assertEqual([], self.__directive_provider.mock_calls)

    def test_process_line_directive(self):
        self.__branch_processor.skip.return_value = True

        l = self.__dp.process_line("# py4lo: ok")
        self.assertEqual([], l)

        self.assertEqual([], self.__scripts_path.mock_calls)
        self.assertEqual([], self.__scripts_processor.mock_calls)
        self.assertEqual([call.handle_directive('ok', [])], self.__branch_processor.mock_calls)
        self.assertEqual([], self.__directive_provider.mock_calls)

    def test_process_line_branch_directive(self):
        self.__branch_processor.skip.return_value = False
        self.__branch_processor.handle_directive.return_value = True

        l = self.__dp.process_line("# py4lo: ok")
        self.assertEqual([], l)

        self.assertEqual([], self.__scripts_path.mock_calls)
        self.assertEqual([], self.__scripts_processor.mock_calls)
        self.assertEqual([call.handle_directive('ok', [])], self.__branch_processor.mock_calls)
        self.assertEqual([], self.__directive_provider.mock_calls)

    def test_process_line_standard_directive(self):
        self.__branch_processor.skip.return_value = False
        self.__branch_processor.handle_directive.return_value = False
        directive = Mock()
        self.__directive_provider.get.return_value = (directive, "")

        l = self.__dp.process_line("# py4lo: ok")
        self.assertEqual([], l)

        self.assertEqual([], self.__scripts_path.mock_calls)
        self.assertEqual([], self.__scripts_processor.mock_calls)
        self.assertEqual([call.handle_directive('ok', []), call.skip()], self.__branch_processor.mock_calls)
        self.assertEqual([call.get(["ok"])], self.__directive_provider.mock_calls)
        self.assertEqual([call.execute(Any(), "")], directive.mock_calls)
