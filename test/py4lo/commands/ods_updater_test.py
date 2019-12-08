#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2019 J. Férard <https://github.com/jferard>
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
from typing import List
from unittest import mock
from unittest.mock import Mock, MagicMock, patch, call

from commands.ods_updater import OdsUpdaterHelper
from core.asset import SourceAsset, DestinationAsset
from core.properties import Sources, Destinations
from core.script import TempScript, SourceScript, DestinationScript


class TestOdsUpdaterHelper(unittest.TestCase):
    def test_assets(self):
        # mocks
        logger: Logger = Mock()
        sources: Sources = Mock()
        destinations: Destinations = Mock()
        source_assets: List[SourceAsset] = Mock()
        destination_asset = DestinationAsset(Path("asset"), bytes())

        # play
        sources.get_assets.return_value = source_assets
        destinations.to_destination_assets.return_value = [destination_asset]

        h = OdsUpdaterHelper(logger, sources, destinations, "3.1")
        assets = h.get_assets()

        # verify
        self.assertEqual([destination_asset], assets)
        self.assertEqual([], logger.mock_calls)
        self.assertEqual([call.get_assets()], sources.mock_calls)
        self.assertEqual([call.to_destination_assets(source_assets)],
                         destinations.mock_calls)

    @patch("commands.ods_updater.ScriptSetProcessor", autospec=True)
    def test_destination_scripts(self, SSPMock):
        # mocks
        logger: Logger = Mock()
        sources: Sources = Mock()
        destinations: Destinations = Mock()
        source_scripts: List[SourceScript] = Mock()
        destination_script: DestinationScript = Mock()
        temp_script: TempScript = Mock()

        # play
        sources.get_src_scripts.return_value = source_scripts
        destinations.to_destination_scripts.return_value = [destination_script]
        SSPMock.return_value.process.return_value = [temp_script]
        h = OdsUpdaterHelper(logger, sources, destinations, "3.1")
        scripts = h.get_destination_scripts()

        # verify
        self.assertEqual([call.debug('Directives tree: %s', mock.ANY)],
                         logger.mock_calls)
        self.assertEqual([destination_script], scripts)
        self.assertEqual([call.get_src_scripts], sources.mock_calls)
        self.assertEqual([call.to_destination_scripts([temp_script])],
                         destinations.mock_calls)