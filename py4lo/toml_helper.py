# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   FastODS is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   FastODS is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""
import toml
import os
import sys
import subprocess
import logging

def load_toml(local_py4lo_toml = "py4lo.toml"):
    py4lo_path = os.path.realpath(os.path.join(os.path.dirname(__file__), ".."))
    return TomlLoader(py4lo_path, local_py4lo_toml).load()

class TomlLoader():
    def __init__(self, py4lo_path, local_py4lo_toml):
        self.__default_py4lo_toml = os.path.join(py4lo_path, "default-py4lo.toml")
        self.__local_py4lo_toml = local_py4lo_toml
        self.__data = { "py4lo_path" : py4lo_path }

    def load(self):
        self.__load_toml()
        self.__check_python_version()
        self.__check_level()
        return self.__data

    def __load_toml(self):
        with open(self.__default_py4lo_toml, 'r', encoding="utf-8") as s:
            content = s.read()
            default_data = toml.loads(content)
            self.__data.update(default_data)

        try:
            with open(self.__local_py4lo_toml, 'r', encoding="utf-8") as s:
                content = s.read()
                local_data = toml.loads(content)
        except OSError as ose:
            pass
        else:
            self.__data.update(local_data)

    def __check_python_version(self):
        if "python_exe" in self.__data:
            status, version = subprocess.getstatusoutput("\""+self.__data["python_exe"]+"\" -V")
            if status == 0:
                self.__data["python_version"] = ((version.split())[1].split("."))[0]

        if not "python_version" in self.__data:
            self.__data["python_exe"] = sys.executable
            self.__data["python_version"] = sys.version_info.major

    def __check_level(self):
        if "log_level" not in self.__data or self.__data["log_level"] not in ["CRITICAL", "DEBUG", "ERROR", "FATAL", "INFO", "NOTSET", "WARN", "WARNING"]:
            self.__data["log_level"] = "INFO"

def __relative_unix_path_to_relative_local_path(path):
    return os.path.sep.join(path.split("/"))