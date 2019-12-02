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
from unittest.mock import Mock, patch, call

from commands.debug_command import *
from py4lo.commands.debug_command import DebugCommand


class TestDebugCommand(unittest.TestCase):
    @patch('zip_updater.ZipUpdater', autospec=True)
    def test(self, Zupdater):
        logger = Mock()
        base_path = Path("")
        src_dir = Path("")
        src_ignore = ["*"]
        assets_dir = Path("")
        assets_ignore = []
        target_dir = Path("")
        assets_dest_dir = Path("")
        python_version = ""
        ods_dest_name = Path("")

        d = DebugCommand(logger, base_path, src_dir, src_ignore, assets_dir,
                         assets_ignore, target_dir, assets_dest_dir,
                         python_version, ods_dest_name)
        d.execute([])

        self.assertEqual([call.info(
            "Debug or init. Generating '%s' for Python '%s'", Path('.'), ''),
            call.log(10, 'Scripts to process: %s', set())],
            logger.mock_calls)
        print(Zupdater.mock_calls)
        self.assertEqual(call().update(Path('inc/debug.ods'), Path('')),
                         Zupdater.mock_calls[-1])
