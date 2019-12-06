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
from logging import Logger
from pathlib import Path
from unittest import mock
from unittest.mock import Mock, call, patch

import env
from commands.ods_updater import OdsUpdaterHelper
from commands.update_command import UpdateCommand
from core.asset import SourceAsset
from core.properties import Destinations, Sources, PropertiesProvider
from core.script import TempScript


class TestUpdateCommand(unittest.TestCase):
    @patch('zip_updater.ZipUpdater', autospec=True)
    def test(self, Zupdater):
        logger: Logger = Mock()
        destinations: Destinations = Mock()
        sources: Sources = Mock()
        asset_path = Mock()
        s = Mock()
        bound = Mock(__enter__=s, __exit__=lambda x, y, z, t: None)
        provider: PropertiesProvider = Mock()

        destinations.dest_ods_file = Path("dest.ods")
        sources.src_dir.rglob.side_effect = [[Path("x.py")]] * 2
        sources.src_ignore = ["*.pyc"]
        sources.source_ods_file = Path("source.ods")
        provider.get_sources.side_effect = [sources]
        provider.get_destinations.side_effect = [destinations]
        provider.get.side_effect = ["3.1", "contact", "3.1"]

        s.read.return_value = "ASSET"
        asset_path.open.return_value = bound

        sources.get_assets.side_effect = [
            [SourceAsset(asset_path, Path("assets"))]]

        directive = UpdateCommand(logger, provider)
        directive.execute(10)

        self.assertEqual([
            call.info("Update. Generating '%s' for Python '%s'",
                      Path('dest.ods'),
                      '3.1'),
            call.debug('Directives tree: %s', mock.ANY),
            call.log(10, 'Scripts to process: %s', []),
            call.info('Updating document %s', Path("dest.ods"))],
            logger.mock_calls)
        self.assertEqual(call().update(Path('source.ods'), Path('dest.ods')),
                         Zupdater.mock_calls[-1])
