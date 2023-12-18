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
import subprocess
import sys
from logging import Logger
from pathlib import Path
from typing import Dict, Callable, Iterator, cast, Optional

from commands.command import Command
from commands.command_executor import CommandExecutor
from core.properties import PropertiesProvider
from core.source_dest import Sources


class TestCommand(Command):
    __test__ = False

    @staticmethod
    def create_executor(
        _args, provider: PropertiesProvider
    ) -> CommandExecutor:
        python_exe = provider.get("python_exe")
        logger = provider.get_logger()
        return CommandExecutor(
            logger, TestCommand(logger, python_exe, provider.get_sources())
        )

    def __init__(self, logger: Logger, python_exe: str, sources: Sources):
        self._logger = logger
        self._python_exe = python_exe
        self._sources = sources
        self._env = cast(Optional[Dict[str, str]], None)

    def execute(self):
        final_status = self._execute_all_tests(
            self._src_paths(), self._execute_doctests
        )
        final_status = (
            self._execute_all_tests(
                self._test_paths(), self._execute_unittests
            )
            or final_status
        )
        return (final_status,)

    def _execute_all_tests(
        self,
        paths: Iterator[Path],
        execute_tests: Callable[[Path], subprocess.CompletedProcess],
    ) -> int:
        final_status = 0
        for path in paths:
            completed_process = execute_tests(path)
            status = completed_process.returncode
            if completed_process.stdout:
                self._logger.info(
                    "output: {0}".format(
                        completed_process.stdout.decode("iso-8859-1")
                    )
                )
            if status != 0:
                if completed_process.stderr:
                    self._logger.error(
                        "error: {0}".format(
                            completed_process.stderr.decode("iso-8859-1")
                        )
                    )
                final_status = 1

        return final_status

    def _execute_unittests(self, path: Path) -> subprocess.CompletedProcess:
        cmd = '"{}" {}'.format(self._python_exe, path)
        self._logger.info("execute unittests: {0}".format(cmd))
        return subprocess.run(
            [self._python_exe, str(path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=self._get_env(),
        )

    def _execute_doctests(self, path: Path) -> subprocess.CompletedProcess:
        cmd = '"{}" -m doctest {}'.format(self._python_exe, path)
        self._logger.info("execute doctests: %s", cmd)
        return subprocess.run(
            [self._python_exe, "-m", "doctest", str(path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=self._get_env(),
        )

    def _test_paths(self) -> Iterator[Path]:
        for path in self._sources.test_dir.rglob("*.py"):
            if path.name.endswith("_test.py") or path.name.startswith("test_"):
                yield path

    def _src_paths(self) -> Iterator[Path]:
        for path in self._sources.src_dir.rglob("*.py"):
            if path != self._sources.src_dir / "main.py":
                yield path

    def _get_env(self) -> Dict[str, str]:
        if self._env is None:
            env = dict(os.environ)
            src_lib = [
                str(self._sources.src_dir),
                str(self._sources.test_dir),
                str(self._sources.opt_dir),
                str(self._sources.lib_dir),
                str(self._sources.inc_dir),
            ]
            env["PYTHONPATH"] = os.pathsep.join(src_lib + sys.path)
            self._env = env
            self._logger.debug("PYTHONPATH = %s", env["PYTHONPATH"])
        return self._env

    @staticmethod
    def get_help():
        return "Do the test of the scripts to add to the spreadsheet"
