# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2017 J. Férard <https://github.com/jferard>

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
import sys
import logging
import subprocess
from commands.command_executor import CommandExecutor

class TestCommand():
    @staticmethod
    def create(args, tdata):
        logger = logging.getLogger("py4lo")
        logger.setLevel(tdata["log_level"])
        return CommandExecutor(TestCommand(logger, tdata["python_exe"], tdata["test_dir"], tdata["src_dir"]))

    def __init__(self, logger, python_exe, test_dir, src_dir):
        self.__logger = logger
        self.__python_exe = python_exe
        self.__test_dir = test_dir
        self.__src_dir = src_dir
        self.__env = None

    def execute(self):
        final_status = 0
        for path in self.__test_paths():
            completed_process = self.__execute(path)
            status = completed_process.returncode
            if completed_process.stdout:
                self.__logger.info("output: {0}".format(completed_process.stdout.decode('ascii')))
            if status != 0:
                if completed_process.stderr:
                    self.__logger.error("error: {0}".format(completed_process.stderr.decode('ascii')))
                final_status = 1

        return (final_status, )

    def __execute(self, path):
        cmd = "\""+self.__python_exe+"\" "+path
        self.__logger.info("execute: {0}".format(cmd))
        return subprocess.run(cmd, capture_output=True, env=self.__get_env())

    def __get_env(self):
        if self.__env is None:
            env = dict(os.environ)
            env["PYTHONPATH"] = ";".join(sys.path+[self.__src_dir])
            self.__env = env
        return self.__env

    def __test_paths(self):
        for dirpath, dirnames, filenames in os.walk(self.__test_dir):
            for filename in filenames:
                if filename.endswith("_test.py"):
                    yield os.path.join(dirpath, filename)

    @staticmethod
    def get_help():
        return "Do the test of the scripts to add to the spreadsheet"
