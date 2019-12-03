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
from asset import Asset
from callbacks import IgnoreItem, RewriteManifest, AddScripts, AddAssets, \
    AddDebugContent
from commands.command import Command, PropertiesProvider
from commands.command_executor import CommandExecutor
from commands.test_command import TestCommand
from script_set_processor import ScriptSetProcessor, TargetScript
from tools import get_assets, get_paths


class DebugCommand(Command):
    @staticmethod
    def create(args, provider: PropertiesProvider):
        test_executor = TestCommand.create(args, provider)
        tdata = provider.get()
        logger = logging.getLogger("py4lo")
        logger.setLevel(tdata["log_level"])
        debug_command = DebugCommand(logger, Path(tdata["base_path"]),
                                     Path(tdata["src_dir"]),
                                     tdata["src_ignore"],
                                     Path(tdata["assets_dir"]),
                                     tdata["assets_ignore"],
                                     Path(tdata["target_dir"]),
                                     Path(tdata["assets_dest_dir"]),
                                     tdata["python_version"],
                                     Path(tdata["debug_file"]))
        return CommandExecutor(debug_command, test_executor)

    def __init__(self, logger: logging.Logger, base_path: Path, src_dir: Path,
                 src_ignore: List[str], assets_dir: Path,
                 assets_ignore: List[str], target_dir: Path,
                 assets_dest_dir: Path, python_version: str,
                 ods_dest_name: Path):
        self._logger = logger
        self._base_path = base_path
        self._src_dir = src_dir
        self._src_ignore = src_ignore
        self._assets_dir = assets_dir
        self._assets_ignore = assets_ignore
        self._target_dir = target_dir
        self._assets_dest_dir = assets_dest_dir
        self._python_version = python_version
        self._ods_dest_name = ods_dest_name
        self._debug_path = target_dir.joinpath(ods_dest_name)

    def execute(self, *_args: List[str]) -> Tuple[Path]:
        self._logger.info("Debug or init. Generating '%s' for Python '%s'",
                          self._debug_path, self._python_version)

        scripts, exported_func_names_by_script = self._get_scripts()
        assets = get_assets(self._assets_dir, self._assets_ignore,
                            self._assets_dest_dir)

        zupdater = self._get_zip_updater(assets, exported_func_names_by_script,
                                         scripts)
        zupdater.update(self._base_path.joinpath("inc", "debug.ods"),
                        self._ods_dest_name)
        return self._ods_dest_name,

    def _get_zip_updater(self, assets: List[Asset],
                         exported_func_names_by_script: Dict[str, List[str]],
                         scripts: List[TargetScript]):
        zupdater_builder = zip_updater.ZipUpdaterBuilder()
        zupdater_builder.item(IgnoreItem(Path("Scripts"))).item(
            RewriteManifest(scripts, assets)).after(AddScripts(scripts)).after(
            AddAssets(assets)).after(
            AddDebugContent(exported_func_names_by_script))
        zupdater = zupdater_builder.build()
        return zupdater

    def _get_scripts(self) -> (List[TargetScript], Dict[str, List[str]]):
        script_paths = get_paths(self._src_dir, self._src_ignore, "*.py")
        scripts_processor = ScriptSetProcessor(self._logger, self._src_dir,
                                               self._target_dir,
                                               self._python_version,
                                               script_paths)
        return scripts_processor.process(), \
               scripts_processor.get_exported_func_names_by_script_name()

    @staticmethod
    def get_help():
        return "Create a debug.ods file with button for each function"
