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
import re
import logging
import py_compile
from directive_processor import DirectiveProcessor

class TargetScript():
    """A target script has a file name, a content (the file was processed),
    some exported functions and the exceptions encountered by the interpreter"""

    @staticmethod
    def to_name(fname):
        return os.path.split(fname)[1]

    def __init__(self, script_fname, script_content, exported_func_names, exception):
        self.__fname = script_fname
        self.__content = script_content
        self.__exported_func_names = exported_func_names
        self.__exception = exception

    def get_fname(self):
        return self.__fname

    def get_name(self):
        return TargetScript.to_name(self.__fname)

    def get_content(self):
        return self.__content

    def get_exported_func_names(self):
        return self.__exported_func_names

    def get_exception(self):
        return self.__exception

class ScriptsProcessor():
    """A script processor processes some file from a source dir and writes
    the target scripts to a target dir"""
    def __init__(self, logger, src_dir, target_dir, python_version):
        self.__logger = logger
        self.__src_dir = src_dir
        self.__target_dir = target_dir
        self.__python_version = python_version

    def process(self, script_fnames):
        """Explore the scripts. Since a script may import another script, we
        have a tree structure. We traverse the tree with classical DFS"""
        self.__logger.log(logging.DEBUG, "Scripts to process: %s", script_fnames)

        self.__scripts = []
        self.__cur_script_fnames = list(script_fnames) # our stack.
        self.__visited = set() # avoid cycles

        while self.__has_more_scripts():
            self.__process_next_script_if_not_visited()

        self.__raise_exceptions()
        return self.__scripts

    def __raise_exceptions(self):
        es = [script.get_exception() for script in self.__scripts if script.get_exception()]
        if not es:
            return

        for e in es:
            self.__logger.log(logging.ERROR, str(e))
        raise Exception("Compilation errors: see above")

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
        directive_processor = DirectiveProcessor.create(self.__src_dir, self, self.__python_version)
        script_processor = ScriptProcessor(self.__logger, directive_processor, script_fname)
        script = script_processor.parse_script()
        self.__add_script(script)

    def __add_script(self, script):
        self.__write_script(script)
        self.__scripts.append(script)

    def get_exported_func_names_by_script_name(self):
        """Return a dict: script name -> exported functions"""
        return {script.get_name():script.get_exported_func_names() for script in self.__scripts}

    def __write_script(self, script):
        script_name = script.get_name()
        self.__ensure_target_dir_exists()
        target_filename = os.path.join(self.__target_dir, script_name)
        self.__logger.log(logging.DEBUG, "Writing script: %s (%s)", script_name, target_filename)
        with open(target_filename, 'wb') as f:
            f.write(script.get_content())

    def __ensure_target_dir_exists(self):
        if not os.path.exists(self.__target_dir):
            os.makedirs(self.__target_dir)

class ScriptProcessor():
    """A script processor"""

    def __init__(self, logger, directive_processor, script_fname):
        self.__logger = logger
        self.__directive_processor = directive_processor
        self.__script_fname = script_fname

    def parse_script(self):
        self.__logger.log(logging.DEBUG, "Parsing script: %s (%s)", TargetScript.to_name(self.__script_fname), self.__script_fname)
        exception = self.__get_exception()
        content, exported_func_names = _ScriptParser(self.__logger, self.__directive_processor, self.__script_fname).parse()
        return TargetScript(self.__script_fname, content.encode("utf-8"), exported_func_names, exception)

    def __get_exception(self):
        try:
            py_compile.compile(self.__script_fname, doraise=True)
        except Exception as e:
            return e
        else:
            return None


class _ScriptParser():
    """A script parser"""

    __PATTERN = re.compile("^def\s+([^_].*?)\(.*\):\s*$")

    def __init__(self, logger, directive_processor, script_fname):
        self.__logger = logger
        self.__directive_processor = directive_processor
        self.__script_fname = script_fname
        self.__script = None

    def parse(self):
        if self.__script:
            return self.__script

        self.__exported_func_names = []
        self.__lines = ["# parsed by py4lo (https://github.com/jferard/py4lo)"]
        with open(self.__script_fname, 'r', encoding="utf-8") as f:
            self.__process_lines(f)

        self.__directive_processor.end()
        self.__add_exported_func_names()
        return "\n".join(self.__lines), self.__exported_func_names

    def __process_lines(self, f):
        try:
            for line in f:
                self.__process_line(line)
        except Exception as e:
            self.__logger.critical(self.__script_fname+", line="+line)
            raise e

    def __process_line(self, line):
        if line[0] == '#':
            self.__lines.extend(self.__directive_processor.process_line(line))
        else:
            self.__append_to_lines(line)

    def __append_to_lines(self, line):
        if self.__directive_processor.ignore_lines():
            self.__lines.append("### py4lo ignore:" + line)
        else:
            self.__update_exported(line)
            self.__lines.append(line)

    def __update_exported(self, line):
        m = _ScriptParser.__PATTERN.match(line)
        if m:
            func_name = m.group(1)
            self.__exported_func_names.append(func_name)

    def __add_exported_func_names(self):
        if self.__exported_func_names:
            self.__lines.extend([
                "",
                "",
                "g_exportedScripts = ({})".format(", ".join(self.__exported_func_names))
            ])
