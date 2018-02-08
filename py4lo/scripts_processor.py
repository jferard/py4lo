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
import re
import logging
import py_compile
from directive_processor import DirectiveProcessor

class Script():
    @staticmethod
    def to_name(fname):
        return os.path.split(fname)[1]

    def __init__(self, script_fname, script_data, exported_func_names, exception):
        self.__fname = script_fname
        self.__data = script_data
        self.__exported_func_names = exported_func_names
        self.__exception = exception

    def get_fname(self):
        return self.__fname

    def get_name(self):
        return Script.to_name(self.__fname)

    def get_data(self):
        return self.__data

    def get_exported_func_names(self):
        return self.__exported_func_names

    def get_exception(self):
        return self.__exception

class ScriptsProcessor():
    def __init__(self, logger, src_dir, python_version, target_dir):
        self.__logger = logger
        self.__src_dir = src_dir
        self.__python_version = python_version
        self.__target_dir = target_dir

    def process(self, script_fnames):
        self.__logger.log(logging.DEBUG, "Scripts to process: %s", script_fnames)

        self.__scripts = []
        self.__cur_script_fnames = list(script_fnames)
        self.__visited = set()

        while self.__has_more_scripts():
            self.__process_next_script_if_not_visited()

        es = [script.get_exception() for script in self.__scripts if script.get_exception()]
        if es:
            for e in es:
                self.__logger.log(logging.ERROR, str(e))
            raise Exception("Compilation errors: see above")

        return self.__scripts

    def append_script(self, script_fname):
        """Append a new script. The directive UseLib will call this method"""
        self.__cur_script_fnames.append(script_fname)

    def __has_more_scripts(self):
        return len(self.__cur_script_fnames)

    def __process_next_script_if_not_visited(self):
        next_script_fname = self.__cur_script_fnames.pop()
        if next_script_fname in self.__visited:
            return # avoid cycles !

        self.__process_script(next_script_fname)
        self.__visited.add(next_script_fname)

    def __process_script(self, script_fname):
        directive_processor = DirectiveProcessor(self, self.__python_version, self.__src_dir)
        script_processor = ScriptProcessor(self.__logger, directive_processor, script_fname)
        script = script_processor.parse_script()
        self.__add_script(script)

    def __add_script(self, script):
        self.__write_script(script)
        self.__scripts.append(script)

    def get_exported_func_names_by_script_name(self):
        """Return a dict: script name -> exported functions"""
        exported_func_names_by_script_name = {}
        for script in self.__scripts:
            exported_func_names_by_script_name[script.get_name()] = script.get_exported_func_names()
        return exported_func_names_by_script_name

    def __write_script(self, script):
        script_name = script.get_name()
        self.__ensure_target_dir_exists()
        target_filename = os.path.join(self.__target_dir, script_name)
        self.__logger.log(logging.DEBUG, "Writing script: %s (%s)", script_name, target_filename)
        with open(target_filename, 'wb') as f:
            f.write(script.get_data())

    def __ensure_target_dir_exists(self):
        if not os.path.exists(self.__target_dir):
            os.mkdir(self.__target_dir)

class ScriptProcessor():
    """A script processor"""
    def __init__(self, logger, directive_processor, script_fname):
        self.__logger = logger
        self.__directive_processor = directive_processor
        self.__script_fname = script_fname

    def parse_script(self):
        self.__logger.log(logging.DEBUG, "Parsing script: %s (%s)", Script.to_name(self.__script_fname), self.__script_fname)
        exported_func_names = []

        exception = None
        try:
            py_compile.compile(self.__script_fname, doraise=True)
        except Exception as e:
            exception = e

        s = "# parsed by py4lo\n"
        with open(self.__script_fname, 'r', encoding="utf-8") as f:
            try:
                for line in f:
                    if line[0] == '#':
                        s += self.__directive_processor.process_line(line)
                    else:
                        # TODO: def without _ -> exported
                        m = re.match("^def\s+(.*?)\(.*\):\s*$", line)
                        if m:
                            func_name = m.group(1)
                            if func_name[0] != "_":
                                exported_func_names.append(func_name)

                        if self.__directive_processor.ignore_lines():
                            s += "### py4lo ignore:" + line
                        else:
                            s += line
            except Exception as e:
                self.__logger.critical(self.__script_fname+", line="+line)
                raise e

        s += "\n\ng_exportedScripts = ("+", ".join(exported_func_names)+")\n"
        script = Script(self.__script_fname, s.encode("utf-8"), exported_func_names, exception)
        self.__directive_processor.end()
        return script
