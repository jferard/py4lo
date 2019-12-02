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
from typing import List
from zipfile import ZipFile

from callbacks import BeforeAfterCallback
from script_set_processor import TargetScript

ARC_SCRIPTS_PATH = "Scripts/python"


class AddScripts(BeforeAfterCallback):
    """After callback. Add some scripts in destination file"""

    def __init__(self, scripts: List[TargetScript]):
        self._scripts = scripts

    def call(self, zout: ZipFile) -> bool:
        for script in self._scripts:
            zout.writestr(ARC_SCRIPTS_PATH + "/" + script.name,
                          script.script_content)
        return True
