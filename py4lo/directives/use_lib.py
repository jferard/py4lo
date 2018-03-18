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

class UseLib():
    sig = "use lib"

    def __init__(self, py4lo_path, scripts_path):
        self.__py4lo_path = py4lo_path

    def execute(self, processor, args):
        """Take a directive processor, a directive name, and args
        The name is used to take or leave this directive. Each directive is responsible for this decision"""
        raise DeprecationWarning("Use 'import lib'")
        (ret, object_ref) = self.__process_args(processor, args)

        processor.include("py4lo_bootstrap.py")
        processor.append(self.__use_lib(ret, object_ref))
        return True

    def __process_args(self, processor, args):
        object_ref = args[0]
        (lib_ref, object_name) = object_ref.split("::")
        script_fname_wo_extension = os.path.join(self.__py4lo_path, "lib", lib_ref)

        script_fname = script_fname_wo_extension + ".py"
        if not os.path.isfile(script_fname):
            raise BaseException(script_fname + " is not a py4lo lib")

        if len(args) == 3 and args[1] == 'as':
            ret = args[2]
        else:
            ret = object_name

        processor.append_script(script_fname)

        return (ret, object_ref)

    def __use_lib(self, ret, object_ref):
        return ret + " = use_local(\""+object_ref+"\")\n"
