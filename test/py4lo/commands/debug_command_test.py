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
import env
import unittest
from unittest.mock import Mock, patch, call

from callbacks import *
from commands.debug_command import *

from py4lo.commands.debug_command import DebugCommand


class TestDebugCommand(unittest.TestCase):
    @patch('zip_updater.ZipUpdater', autospec=True)
    @patch('os.listdir')
    def test(self, listdir, Zupdater):
        logger = Mock()
        py4lo_path = ""
        src_dir = ""
        assets_dir = ""
        target_dir = ""
        assets_dest_dir = ""
        python_version = ""
        ods_dest_name = ""

        listdir.return_value = []

        d = DebugCommand(logger, py4lo_path, src_dir, assets_dir, target_dir, assets_dest_dir, python_version, ods_dest_name)
        d.execute([])

        self.assertEqual([call.info("Debug or init. Generating '%s' for Python '%s'", '', ''),
                          call.log(10, 'Scripts to process: %s', set())], logger.mock_calls)
        self.assertEqual([call('')], listdir.mock_calls)
        self.assertEqual(call(), Zupdater.mock_calls[0])
        self.assertEqual(call().update('inc/debug.ods', ''), Zupdater.mock_calls[-1])
