# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
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
import unittest
from logging import Logger
from pathlib import Path
from unittest import mock

from commands.debug_command import DebugCommand
from commands.ods_updater import OdsUpdaterHelper
from core.source_dest import Sources, Destinations
from core.script import TempScript


class TestDebugCommand(unittest.TestCase):
    @mock.patch('zip_updater.ZipUpdater', autospec=True)
    def test(self, Zupdater):
        # mocks
        logger: Logger = mock.Mock()
        helper: OdsUpdaterHelper = mock.Mock()
        t1: TempScript = mock.Mock()
        t2: TempScript = mock.Mock()
        sources: Sources = mock.Mock()
        destinations: Destinations = mock.Mock()

        dest_path = Path("dest.ods")
        inc_path = Path("inc/debug.ods")
        python_version = "3.1"
        helper.get_temp_scripts.side_effect = [[t1, t2]]
        destinations.dest_ods_file.parent.joinpath.return_value = dest_path
        sources.inc_dir.joinpath.return_value = inc_path
        d = DebugCommand(logger, helper, sources, destinations, python_version)

        d.execute([])

        self.assertEqual([mock.call.info(
            "Debug or init. Generating '%s' for Python '%s'", Path('dest.ods'),
            '3.1')],
            logger.mock_calls)
        print(Zupdater.mock_calls)
        self.assertEqual(mock.call().update(Path('inc/debug.ods'),
                                            Path('dest.ods')),
                         Zupdater.mock_calls[-1])
