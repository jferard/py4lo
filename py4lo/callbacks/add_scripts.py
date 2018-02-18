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

ARC_SCRIPTS_PATH = "Scripts/python"

class AddScripts():
    """After callback. Add some scripts in destination file"""
    def __init__(self, scripts):
        self.__scripts = scripts

    def call(self, zout):
        for script in self.__scripts:
            zout.writestr(ARC_SCRIPTS_PATH+"/"+script.get_name(), script.get_data())
        return True
