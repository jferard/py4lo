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
import logging
from logging import Logger
import os
import subprocess
import sys
from pathlib import Path
from typing import Dict, Callable, Iterator

from commands.command import Command, PropertiesProvider
from commands.command_executor import CommandExecutor


class TestCommand(Command):
    __test__ = False

    @staticmethod
    def create_executor(_args, provider: PropertiesProvider):
        tdata = provider.get()
        logger = logging.getLogger("py4lo")
        logger.setLevel(tdata["log_level"])
        return CommandExecutor(
            TestCommand(logger, tdata["python_exe"], Path(tdata["test_dir"]),
                        Path(tdata["src_dir"]), Path(tdata["base_path"])))

    def __init__(self, logger: Logger, python_exe: str, test_dir: Path,
                 src_dir: Path, base_path: Path):
        self._logger = logger
        self._python_exe = python_exe
        self._test_dir = test_dir
        self._src_dir = src_dir
        self._base_path = base_path
        self._env = None

    def execute(self):
        final_status = self._execute_all_tests(self._src_paths(),
                                               self._execute_doctests)
        final_status = self._execute_all_tests(self._test_paths(),
                                               self._execute_unittests) or final_status
        return final_status,

    def _execute_all_tests(self, paths: Iterator[Path],
                           execute_tests: Callable[
                               [Path], subprocess.CompletedProcess]) -> int:
        final_status = 0
        for path in paths:
            completed_process = execute_tests(path)
            status = completed_process.returncode
            if completed_process.stdout:
                self._logger.info("output: {0}".format(
                    completed_process.stdout.decode('iso-8859-1')))
            if status != 0:
                if completed_process.stderr:
                    self._logger.error("error: {0}".format(
                        completed_process.stderr.decode('iso-8859-1')))
                final_status = 1

        return final_status

    def _execute_unittests(self, path: Path) -> subprocess.CompletedProcess:
        cmd = "\"{}\" {}".format(self._python_exe, path)
        self._logger.info("execute: {0}".format(cmd))
        return subprocess.run([self._python_exe, str(path)],
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              env=self._get_env())

    def _execute_doctests(self, path: Path) -> subprocess.CompletedProcess:
        cmd = "\"{}\" -m doctest {}".format(self._python_exe, path)
        self._logger.info("execute: %s", cmd)
        return subprocess.run([self._python_exe, "-m", "doctest", str(path)],
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              env=self._get_env())

    def _test_paths(self) -> Iterator[Path]:
        for path in self._test_dir.rglob("*.py"):
            if path.name.endswith("_test.py"):
                yield path

    def _src_paths(self) -> Iterator[Path]:
        for path in self._src_dir.rglob("*.py"):
            if path.name != "main.py":
                yield path

    def _get_env(self) -> Dict[str, str]:
        if self._env is None:
            env = dict(os.environ)
            src_lib = [str(self._src_dir),
                       str(self._base_path.joinpath("lib"))]
            env["PYTHONPATH"] = ";".join(sys.path + src_lib)
            self._env = env
        return self._env

    @staticmethod
    def get_help():
        return "Do the test of the scripts to add to the spreadsheet"
