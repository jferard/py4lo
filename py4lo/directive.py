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
import shlex
import os

class UseLib():
    def __init__(self, py4lo_path):
        self.__py4lo_path = py4lo_path

    def execute(self, processor, directiveName, args):
        if directiveName != "use" or args[0] != "lib":
            return False

        (ret, object_ref) = self.__process_args(processor, args[1:])

        processor.bootstrap()
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

class UseObject():
    def __init__(self, scripts_path):
        self.__scripts_path = scripts_path

    def execute(self, processor, directiveName, args):
        if directiveName != "use" or args[0] == "lib":
            return False

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

class Include():
    def __init__(self, scripts_path):
        self.__scripts_path = scripts_path

    def execute(self, processor, directiveName, args):
        if directiveName != "include":
            return False

        filename = os.path.join(self.__scripts_path, "inc", args[0]+".py")

        s = "# begin py4lo include: "+filename+"\n"
        with open(filename, 'r', encoding='utf-8') as f:
            for line in f:
                s += line
        s += "\n# end py4lo include\n"

        processor.append(s)
        return True

class ImportLib():
    def __init__(self, py4lo_path):
        self.__py4lo_path = py4lo_path

    def execute(self, processor, directiveName, args):
        if directiveName != "import" or args[0] != "lib":
            return False

        processor.import2()
        script_ref = args[1]
        script_fname = os.path.join(self.__py4lo_path, "lib", script_ref+".py")
        print (script_fname)
        processor.append_script(script_fname)
        processor.append("import "+script_ref+"\n")
        processor.append("try:\n")
        processor.append("    "+script_ref+".init(XSCRIPTCONTEXT)\n")
        processor.append("except NameError:\n")
        processor.append("    pass\n")
        return True

class Import():
    def __init__(self, scripts_path):
        self.__scripts_path = scripts_path

    def execute(self, processor, directiveName, args):
        if directiveName != "import" or args[0] == "lib":
            return False

        processor.import2()
        script_ref = args[0]
        script_fname = os.path.join(self.__scripts_path, script_ref+".py")
        processor.append_script(script_fname)
        processor.append("import "+script_ref+"\n")
        return True

class Fail():
    def execute(self, processor, directiveName, args):
        print("Wrong directive "+directive+" (line ="+line)
