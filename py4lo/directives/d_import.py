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

class Import():
    sig = "import"

    def __init__(self, py4lo_path, scripts_path):
        self.__scripts_path = scripts_path

    def execute(self, processor, args):
        script_ref = args[0]
        processor.import2()
        script_fname = os.path.join(self.__scripts_path, script_ref+".py")
        processor.append_script(script_fname)
        processor.append("import "+script_ref+"\n")
        return True
