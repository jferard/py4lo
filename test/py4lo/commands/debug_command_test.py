# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2021 J. Férard <https://github.com/jferard>

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
from logging import Logger
from unittest.mock import Mock, patch, call

from commands.debug_command import *
from commands.debug_command import DebugCommand
from core.source_dest import Sources, Destinations
from core.script import TempScript


class TestDebugCommand(unittest.TestCase):
    @patch('zip_updater.ZipUpdater', autospec=True)
    def test(self, Zupdater):
        # mocks
        logger: Logger = Mock()
        helper: OdsUpdaterHelper = Mock()
        t1: TempScript = Mock()
        t2: TempScript = Mock()
        sources: Sources = Mock()
        destinations: Destinations = Mock()

        dest_path = Path("dest.ods")
        inc_path = Path("inc/debug.ods")
        python_version = "3.1"
        helper.get_temp_scripts.side_effect = [[t1, t2]]
        destinations.dest_ods_file.parent.joinpath.return_value = dest_path
        sources.inc_dir.joinpath.return_value = inc_path
        d = DebugCommand(logger, helper, sources, destinations, python_version)

        d.execute([])

        self.assertEqual([call.info(
            "Debug or init. Generating '%s' for Python '%s'", Path('dest.ods'),
            '3.1')],
            logger.mock_calls)
        print(Zupdater.mock_calls)
        self.assertEqual(call().update(Path('inc/debug.ods'), Path('dest.ods')),
                         Zupdater.mock_calls[-1])
