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
import logging
from pathlib import Path
from typing import Optional, List

from callbacks import AddReadmeWith, IgnoreItem, ARC_SCRIPTS_PATH, \
    RewriteManifest, AddScripts, AddAssets
from commands.command import Command
from commands.command_executor import CommandExecutor
from commands.ods_updater import OdsUpdaterHelper
from commands.test_command import TestCommand
from core.asset import DestinationAsset
from core.properties import PropertiesProvider, Sources, Destinations
from core.script import DestinationScript
from zip_updater import ZipUpdater, ZipUpdaterBuilder


class UpdateCommand(Command):
    @staticmethod
    def create_executor(args, provider: PropertiesProvider):
        test_executor = TestCommand.create_executor(args, provider)
        update_command = UpdateCommand(provider)
        return CommandExecutor(update_command, test_executor)

    def __init__(self, provider: PropertiesProvider):
        self._provider = provider

    def execute(self, status: int) -> (int, Path):
        sources = self._provider.get_sources()
        destinations = self._provider.get_destinations()
        dest_name = _UpdateCommandHelper.create(self._provider, sources,
                                                destinations).update(
            sources.source_ods_file, destinations.dest_ods_file)
        return status, dest_name

    @staticmethod
    def get_help():
        return "Update the file with all scripts"


class _UpdateCommandHelper:
    @staticmethod
    def create(provider: PropertiesProvider, sources: Sources,
               destinations: Destinations) -> "_UpdateCommandHelper":
        logger = provider.get_logger()
        tdata = provider.get()
        add_readme = tdata.get("add_readme", False)
        if add_readme:
            readme_contact = tdata["readme_contact"]
            add_readme_callback = AddReadmeWith(provider.get_base_path().joinpath("inc"),
                                                readme_contact)
        else:
            add_readme_callback = None

        python_version = tdata["python_version"]
        helper = OdsUpdaterHelper(logger, sources, destinations, python_version)

        return _UpdateCommandHelper(logger, helper, destinations,
                                    python_version,
                                    add_readme_callback)

    def __init__(self, logger: logging.Logger, helper: OdsUpdaterHelper,
                 destinations: Destinations, python_version: str,
                 add_readme_callback: Optional[AddReadmeWith]):
        self._helper = helper
        self._logger = logger
        self._destinations = destinations
        self._python_version = python_version
        self._add_readme_callback = add_readme_callback

    def update(self, ods_source: Path, ods_dest: Path) -> Path:
        self._logger.info("Debug or init. Generating %s for Python %s",
                          ods_dest, self._python_version)

        temp_scripts = self._helper.get_temp_scripts()
        scripts = [ts.to_destination(self._destinations.dest_dir) for ts in
                   temp_scripts]
        assets = self._helper.get_assets()

        self._logger.info("Updating document %s",
                          ods_dest)
        zip_updater = self._create_updater(scripts, assets)
        zip_updater.update(ods_source, ods_dest)
        return ods_dest

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
