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
from typing import List, Tuple, Dict

import zip_updater
from callbacks import IgnoreItem, RewriteManifest, AddScripts, AddAssets, \
    AddDebugContent
from commands.command import Command
from commands.command_executor import CommandExecutor
from commands.ods_updater import OdsUpdaterHelper
from commands.test_command import TestCommand
from core.asset import DestinationAsset
from core.properties import PropertiesProvider, Destinations, Sources
from core.script import DestinationScript


class DebugCommand(Command):
    """
    Create a simple ods file with a button per exported function.
    This is a good start.
    """

    @staticmethod
    def create_executor(args, provider: PropertiesProvider):
        test_executor = TestCommand.create_executor(args, provider)
        tdata = provider.get()
        logger = logging.getLogger("py4lo")
        logger.setLevel(tdata["log_level"])
        sources = provider.get_sources()
        destinations = provider.get_destinations()
        debug_command = DebugCommand(logger, Path(tdata["base_path"]), sources,
                                     tdata["python_version"],
                                     Path(tdata["debug_file"]))
        return CommandExecutor(debug_command, test_executor)

    def __init__(self, logger: logging.Logger, helper: OdsUpdaterHelper,
                 sources: Sources, destinations: Destinations,
                 python_version: str):
        self._logger = logger
        self._helper = helper
        self._sources = sources
        self._destinations = destinations
        self._python_version = python_version
        self._debug_path = destinations.target_dir.joinpath(
            destinations.dest_ods_file)

    def execute(self, *_args: List[str]) -> Tuple[Path]:
        self._logger.info("Debug or init. Generating '%s' for Python '%s'",
                          self._debug_path, self._python_version)

        temp_scripts = self._helper.get_temp_scripts()
        exported_func_names_by_script = {ts.dest_name: ts.exported_func_names
                                         for ts in temp_scripts}
        dest_scripts = [ts.to_destination(self._destinations.dest_dir) for ts in
                        temp_scripts]
        assets = self._helper.get_assets()

        zupdater = self._get_zip_updater(assets, exported_func_names_by_script,
                                         dest_scripts)
        zupdater.update(self._sources.inc_dir.joinpath("debug.ods"),
                        self._debug_path)
        return self._debug_path,

    def _get_zip_updater(self, assets: List[DestinationAsset],
                         exported_func_names_by_script: Dict[str, List[str]],
                         scripts: List[DestinationScript]):
        zupdater_builder = zip_updater.ZipUpdaterBuilder(self._logger)
        zupdater_builder.item(IgnoreItem(Path("Scripts"))).item(
            RewriteManifest(scripts, assets)).after(
            AddScripts(self._logger, scripts)).after(
            AddAssets(assets)).after(
            AddDebugContent(exported_func_names_by_script))
        zupdater = zupdater_builder.build()
        return zupdater

    @staticmethod
    def get_help():
        return "Create a debug.ods file with button for each function"
