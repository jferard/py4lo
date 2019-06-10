# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2018 J. Férard <https://github.com/jferard>

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

import os
from commands.test_command import TestCommand
from commands.command_executor import CommandExecutor
import logging
from scripts_processor import ScriptsProcessor
import scripts_processor
from callbacks import *
import zip_updater


class DebugCommand:
    @staticmethod
    def create(args, tdata):
        test_executor = TestCommand.create(args, tdata)
        logger = logging.getLogger("py4lo")
        logger.setLevel(tdata["log_level"])
        debug_command = DebugCommand(logger, tdata["py4lo_path"], tdata["src_dir"], tdata["assets_dir"],
                                         tdata["target_dir"], tdata["assets_dest_dir"], tdata["python_version"],
                                         tdata["debug_file"])
        return CommandExecutor(debug_command, test_executor)

    def __init__(self, logger, py4lo_path, src_dir, assets_dir, target_dir, assets_dest_dir, python_version,
                 ods_dest_name):
        self._logger = logger
        self._py4lo_path = py4lo_path
        self._src_dir = src_dir
        self._assets_dir = assets_dir
        self._target_dir = target_dir
        self._assets_dest_dir = assets_dest_dir
        self._python_version = python_version
        self._ods_dest_name = ods_dest_name
        self._debug_path = os.path.join(target_dir, ods_dest_name)

    def execute(self, *_args):
        self._logger.info("Debug or init. Generating '%s' for Python '%s'", self._debug_path, self._python_version)

        scripts_processor = ScriptsProcessor(self._logger, self._src_dir, self._target_dir, self._python_version)
        scripts = self._get_scripts(scripts_processor)
        assets = self._get_assets()

        zupdater = zip_updater.ZipUpdater()
        (
            zupdater
                .item(IgnoreScripts("Scripts"))
                .item(RewriteManifest(scripts, assets))
                .after(AddScripts(scripts))
                .after(AddAssets(assets))
                .after(AddDebugContent(scripts_processor.get_exported_func_names_by_script_name()))
        )
        zupdater.update(os.path.join(self._py4lo_path, "inc", "debug.ods"), self._ods_dest_name)
        return self._ods_dest_name,

    def _get_scripts(self, scripts_processor):
        script_fnames = set(
            os.path.join(self._src_dir, fname) for fname in os.listdir(self._src_dir) if fname.endswith(".py"))
        return scripts_processor.process(script_fnames)

    def _get_assets(self):
        assets = []
        for root, _, fnames in os.walk(self._assets_dir):
            for fname in fnames:
                filename = os.path.join(root, fname)
                dest_name = os.path.join(self._assets_dest_dir, os.path.relpath(filename, self._assets_dir)).replace(
                    os.path.sep, "/")
                with open(filename, 'rb') as source:
                    assets.append(Asset(dest_name, source.read()))

        return assets

    @staticmethod
    def get_help():
        return "Create a debug.ods file with button for each function"
