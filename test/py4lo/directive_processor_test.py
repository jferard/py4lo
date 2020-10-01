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
from pathlib import Path
from unittest import mock
from unittest.mock import *

import tst_env
from directive_processor import *
from script_set_processor import ScriptProcessor


class TestDirectiveProcessor(unittest.TestCase):
    def setUp(self):
        self._scripts_path: Path = Mock()
        self._scripts_processor: ScriptProcessor = Mock()
        self._branch_processor: BranchProcessor = Mock()
        self._directive_provider: DirectiveProvider = Mock()
        self._dp = DirectiveProcessor(self._scripts_processor,
                                      self._branch_processor,
                                      self._directive_provider)

    def test_create(self):
        script = SourceScript(Path("ok"), Path("."))
        dp = DirectiveProcessor.create(self._scripts_processor,
                                       self._directive_provider, "3.6", script)

    def test_append_script(self):
        script = SourceScript(Path("ok"), Path("."))
        self._dp.append_script(script)

        self.assertEqual([], self._scripts_path.mock_calls)
        self.assertEqual([call.append_script(script)],
                         self._scripts_processor.mock_calls)
        self.assertEqual([], self._branch_processor.mock_calls)
        self.assertEqual([], self._directive_provider.mock_calls)

    def test_process_line_comment(self):
        self._branch_processor.skip.return_value = True

        l = self._dp.process_line("ok")
        self.assertEqual(["### ok"], l)

        self.assertEqual([], self._scripts_path.mock_calls)
        self.assertEqual([], self._scripts_processor.mock_calls)
        self.assertEqual([call.skip()], self._branch_processor.mock_calls)
        self.assertEqual([], self._directive_provider.mock_calls)

    def test_process_line_write(self):
        self._branch_processor.skip.return_value = False

        l = self._dp.process_line("ok")
        self.assertEqual(["ok"], l)

        self.assertEqual([], self._scripts_path.mock_calls)
        self.assertEqual([], self._scripts_processor.mock_calls)
        self.assertEqual([call.skip()], self._branch_processor.mock_calls)
        self.assertEqual([], self._directive_provider.mock_calls)

    def test_process_line_directive(self):
        self._branch_processor.skip.return_value = True

        l = self._dp.process_line("# py4lo: ok")
        self.assertEqual([], l)

        self.assertEqual([], self._scripts_path.mock_calls)
        self.assertEqual([], self._scripts_processor.mock_calls)
        self.assertEqual([call.handle_directive('ok', [])],
                         self._branch_processor.mock_calls)
        self.assertEqual([], self._directive_provider.mock_calls)

    def test_process_line_branch_directive(self):
        self._branch_processor.skip.return_value = False
        self._branch_processor.handle_directive.return_value = True

        l = self._dp.process_line("# py4lo: ok")
        self.assertEqual([], l)

        self.assertEqual([], self._scripts_path.mock_calls)
        self.assertEqual([], self._scripts_processor.mock_calls)
        self.assertEqual([call.handle_directive('ok', [])],
                         self._branch_processor.mock_calls)
        self.assertEqual([], self._directive_provider.mock_calls)

    def test_process_line_standard_directive(self):
        self._branch_processor.skip.return_value = False
        self._branch_processor.handle_directive.return_value = False
        directive = Mock()
        self._directive_provider.get.return_value = (directive, [""])

        l = self._dp.process_line("# py4lo: ok")
        self.assertEqual([], l)

        self.assertEqual([], self._scripts_path.mock_calls)
        self.assertEqual([], self._scripts_processor.mock_calls)
        self.assertEqual([call.handle_directive('ok', []), call.skip()],
                         self._branch_processor.mock_calls)
        self.assertEqual([call.get(["ok"])],
                         self._directive_provider.mock_calls)
        self.assertEqual([call.execute(self._dp, mock.ANY, [""])],
                         directive.mock_calls)

    # def test_include(self):
    #     lines = self._dp.include("py4lo_import.py")
    #     self.assertEqual(['try:',
    #                       '    XSCRIPTCONTEXT',
    #                       'except:',
    #                       '    pass',
    #                       'else:',
    #                       '    import unohelper',
    #                       '    import sys',
    #                       '    doc = XSCRIPTCONTEXT.getDocument()',
    #                       "    spath = unohelper.fileUrlToSystemPath(doc.URL+'/Scripts/python')",
    #                       '    if spath not in sys.path:',
    #                       '        sys.path.insert(0, spath)'], lines)
    #
    #     self.assertEqual([], self._scripts_path.mock_calls)
    #     self.assertEqual([], self._scripts_processor.mock_calls)
    #     self.assertEqual([], self._branch_processor.mock_calls)
    #     self.assertEqual([], self._directive_provider.mock_calls)
