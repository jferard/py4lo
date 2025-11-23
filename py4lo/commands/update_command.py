# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2025 J. Férard <https://github.com/jferard>
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
from logging import Logger
from pathlib import Path
from typing import Optional, List, Tuple, Any

from callbacks import (AddReadmeWith, IgnoreItem, ARC_SCRIPTS_PATH,
                       RewriteManifest, AddScripts, AddAssets)
from commands.command import Command
from commands.command_executor import CommandExecutor
from commands.ods_updater import OdsUpdaterHelper
from commands.test_command import TestCommand
from core.asset import DestinationAsset
from core.properties import PropertiesProvider
from core.script import DestinationScript
from zip_updater import ZipUpdater, ZipUpdaterBuilder


class UpdateCommand(Command):
    @staticmethod
    def create_executor(args: List[str], provider: PropertiesProvider) -> CommandExecutor:
        if "notest" in args:
            test_executor = None
        else:
            test_executor = TestCommand.create_executor(args, provider)
        logger = provider.get_logger()
        update_command = UpdateCommand.create(logger, provider)
        return CommandExecutor(logger, update_command, test_executor)

    @staticmethod
    def create(logger: Logger,
               provider: PropertiesProvider) -> Command:
        sources = provider.get_sources()
        destinations = provider.get_destinations()
        add_readme_callback = provider.get_readme_callback()
        python_version = provider.get("python_version")
        source_ods_file = sources.source_ods_file
        dest_ods_file = destinations.dest_ods_file
        helper = OdsUpdaterHelper(
            logger, sources, destinations, python_version)
        return UpdateCommand(logger, helper, source_ods_file, dest_ods_file,
                             python_version, add_readme_callback)

    def __init__(self, logger: Logger, helper: OdsUpdaterHelper,
                 source_ods_file: Path,
                 dest_ods_file: Path, python_version: str,
                 add_readme_callback: Optional[AddReadmeWith]):
        self._logger = logger
        self._helper = helper
        self._source_ods_file = source_ods_file
        self._dest_ods_file = dest_ods_file
        self._python_version = python_version
        self._add_readme_callback = add_readme_callback

    def execute(self, status: int = 0) -> Tuple[Any, ...]:
        self._logger.info(
            "Update. Generating '%s' (source: %s) for Python '%s'",
            self._dest_ods_file,
            self._source_ods_file,
            self._python_version)
        scripts = self._helper.get_destination_scripts()
        assets = self._helper.get_assets()
        zip_updater = self._create_updater(scripts, assets)
        zip_updater.update(self._source_ods_file,
                           self._dest_ods_file)
        return status, self._dest_ods_file

    def _create_updater(self, scripts: List[DestinationScript],
                        assets: List[DestinationAsset]) -> ZipUpdater:
        zip_updater_builder = ZipUpdaterBuilder(self._logger)
        (
            zip_updater_builder.item(IgnoreItem(ARC_SCRIPTS_PATH))
            .item(RewriteManifest(scripts, assets))
            .after(AddScripts(self._logger, scripts))
            .after(AddAssets(assets))
        )
        if self._add_readme_callback is not None:
            zip_updater_builder.after(self._add_readme_callback)

        return zip_updater_builder.build()

    def get_help(self) -> str:
        return "Update the file with all scripts"
