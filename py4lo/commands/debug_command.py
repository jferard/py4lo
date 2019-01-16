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
from callbacks import *
import zip_updater

class DebugCommand:
    @staticmethod
    def create(args, tdata):
        test_executor = TestCommand.create(args, tdata)
        logger = logging.getLogger("py4lo")
        logger.setLevel(tdata["log_level"])
        debug_command = DebugCommand(logger, tdata["py4lo_path"], tdata["src_dir"], tdata["target_dir"], tdata["python_version"], tdata["debug_file"])
        return CommandExecutor(debug_command, test_executor)

    def __init__(self, logger, py4lo_path, src_dir, target_dir, python_version, ods_dest_name):
        self.__logger = logger
        self.__py4lo_path = py4lo_path
        self.__src_dir = src_dir
        self.__target_dir = target_dir
        self.__python_version = python_version
        self.__ods_dest_name = ods_dest_name
        self.__debug_path = os.path.join(target_dir, ods_dest_name)

    def execute(self, *_args):
        self.__logger.info("Debug or init. Generating '%s' for Python '%s'", self.__debug_path, self.__python_version)

        script_fnames = set(os.path.join(self.__src_dir, fname) for fname in os.listdir(self.__src_dir) if fname.endswith(".py"))
        scripts_processor = ScriptsProcessor(self.__logger, self.__src_dir, self.__target_dir, self.__python_version)
        scripts = scripts_processor.process(script_fnames)
        zupdater = zip_updater.ZipUpdater()
        (
            zupdater
            .item(IgnoreScripts("Scripts"))
            .item(RewriteManifest(scripts))
            .after(AddScripts(scripts))
            .after(AddDebugContent(scripts_processor.get_exported_func_names_by_script_name()))
        )
        zupdater.update(os.path.join(self.__py4lo_path, "inc", "debug.ods"), self.__ods_dest_name)
        return self.__ods_dest_name,

    @staticmethod
    def get_help():
        return "Create a debug.ods file with button for each function"
