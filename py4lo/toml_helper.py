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
import subprocess
import sys
from pathlib import Path
from typing import Dict

import toml


class TomlLoader:
    """Load a toml file and merge values with the default toml file"""

    def __init__(self, py4lo_path: Path, local_py4lo_toml: Path):
        self._default_py4lo_toml = py4lo_path.joinpath(
            "default-py4lo.toml")
        self._local_py4lo_toml = local_py4lo_toml
        self._data: Dict[str, object] = {"py4lo_path": py4lo_path}

    def load(self) -> Dict[str, object]:
        self._load_toml(self._default_py4lo_toml)
        self._load_toml(self._local_py4lo_toml)
        self._check_python_target_version()
        self._check_level()
        return self._data

    def _load_toml(self, path: Path):
        try:
            with path.open('r', encoding="utf-8") as s:
                content = s.read()
                data = toml.loads(content)
        except Exception as e:
            print("Error when loading toml file {}: {}".format(path, e))
        else:
            self._data.update(data)

    def _check_python_target_version(self):
        # get version from target executable
        if "python_exe" in self._data:
            status, version = subprocess.getstatusoutput(
                "\"" + str(self._data["python_exe"]) + "\" -V")
            if status == 0:
                self._data["python_version"] = \
                    ((version.split())[1].split("."))[0]
                return

        # if python_exe was not set, or did not return the expected result,
        # get from sys. It's the local python.
        if "python_version" not in self._data:
            self._data["python_exe"] = sys.executable
            self._data["python_version"] = str(
                sys.version_info.major) + "." + str(sys.version_info.minor)

    def _check_level(self):
        if "log_level" not in self._data or self._data["log_level"] not in [
            "CRITICAL", "DEBUG", "ERROR", "FATAL",
            "INFO", "NOTSET", "WARN", "WARNING"]:
            self._data["log_level"] = "INFO"


def load_toml(local_py4lo_toml: Path = Path("py4lo.toml")) -> Dict[str, object]:
    py4lo_path = Path(__file__).parent
    return TomlLoader(py4lo_path, local_py4lo_toml).load()
