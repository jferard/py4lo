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

class UseObject():
    sig = "use object"

    def __init__(self, py4lo_path, scripts_path):
        self.__scripts_path = scripts_path

    def execute(self, processor, args):
        raise DeprecationWarning("Use 'import lib'")
        processor.bootstrap()
        processor.append(self.__use_object(processor, args))
        return True

    def __use_object(self, processor, args):
        object_ref = args[0]
        (script_ref, object_name) = object_ref.split("::")
        script_fname = os.path.join(self.__scripts_path, "", script_ref+".py")
        if not os.path.isfile(script_fname):
            raise BaseException(script_fname + " is not a script")

        if len(args) == 3 and args[1] == 'as':
            ret = args[2]
        else:
            ret = object_name

        processor.append_script(script_fname)
        return ret + " = use_local(\""+object_ref+"\")\n"
