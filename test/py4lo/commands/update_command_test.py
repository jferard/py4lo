# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
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

from commands.ods_updater import OdsUpdaterHelper
from commands.update_command import UpdateCommand
from zip_updater import ZipUpdaterBuilder


class TestUpdateCommand(unittest.TestCase):
    @mock.patch("commands.update_command.ZipUpdaterBuilder", autospec=True)
    def test_execute(self, ZubMock):
        # mocks
        logger: Logger = mock.Mock()
        helper: OdsUpdaterHelper = mock.Mock()
        zub: ZipUpdaterBuilder = mock.Mock()

        ZubMock.return_value = zub

        # play
        directive = UpdateCommand(
            logger, helper, Path("source.ods"), Path("dest.ods"), "3.1", None
        )
        execute = directive.execute(10)

        # verify
        self.assertEqual((10, Path("dest.ods")), execute)
        self.assertEqual(
            [
                mock.call.info(
                    "Update. Generating '%s' (source: %s) for Python '%s'",
                    Path("dest.ods"),
                    Path("source.ods"),
                    "3.1",
                )
            ],
            logger.mock_calls,
        )
        self.assertEqual(
            [mock.call.get_destination_scripts(), mock.call.get_assets()],
            helper.mock_calls,
        )
        self.assertEqual(
            mock.call.build().update(Path("source.ods"), Path("dest.ods")),
            zub.mock_calls[-1],
        )

    # @patch('zip_updater.ZipUpdaterBuilder', autospec=True)
    # def test_helper(self, ZupdaterB):
    #     logger: Logger = Mock()
    #     helper: OdsUpdaterHelper = Mock()
    #     zup: ZipU
    #
    #     ZupdaterB.return_value = zup
    #
    #     h = _UpdateCommandHelper(logger, helper, "3.1", None)
    #     h.update(Path("source.ods"), Path("dest.ods"))
    #
    #     self.assertEqual([call.get_destination_scripts(), call.get_assets()],
    #                      helper.mock_calls)
    #     self.assertEqual([], ZupdaterB.mock_calls)

    # @patch('zip_updater.ZipUpdater', autospec=True)
    # def test_helper(self, Zupdater):
    #     logger: Logger = Mock()
    #     destinations: Destinations = Mock()
    #     sources: Sources = Mock()
    #     asset_path = Mock()
    #     s = Mock()
    #     bound = Mock(__enter__=s, __exit__=lambda x, y, z, t: None)
    #     provider: PropertiesProvider = Mock()
    #
    #     destinations.dest_ods_file = Path("dest.ods")
    #     sources.get_src_paths.side_effect = [[Path("x.py")]]
    #     sources.src_dir = Path(".")
    #     #        sources.relative_path = PropertyMock(return_value=Path("x.py"))
    #     sources.source_ods_file = Path("source.ods")
    #     provider.get_sources.side_effect = [sources]
    #     provider.get_destinations.side_effect = [destinations]
    #     provider.get.side_effect = ["3.1", "contact", "3.1"]
    #
    #     s.read.return_value = "ASSET"
    #     asset_path.open.return_value = bound
    #
    #     sources.get_assets.side_effect = [
    #         [SourceAsset(asset_path, Path("assets"))]]
    #
    #     directive = _UpdateCommandHelper(logger, helper, "3.1", destinations)
    #     directive.execute(10)
    #
    #     self.assertEqual([
    #         call.info("Update. Generating '%s' for Python '%s'",
    #                   Path('dest.ods'),
    #                   '3.1'),
    #         call.debug('Directives tree: %s', mock.ANY),
    #         call.debug('Scripts to process: %s', []),
    #         call.info('Updating document %s', Path("dest.ods"))],
    #         logger.mock_calls)
    #     self.assertEqual(call().update(Path('source.ods'), Path('dest.ods')),
    #                      Zupdater.mock_calls[-1])
