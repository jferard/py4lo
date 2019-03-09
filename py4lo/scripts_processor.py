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


class TargetScript:
    """A target script has a file name, a content (the file was processed),
    some exported functions and the exceptions encountered by the interpreter"""

    @staticmethod
    def to_name(fname):
        return os.path.split(fname)[1]

    def __init__(self, script_fname, script_content, exported_func_names, exception):
        self._fname = script_fname
        self._content = script_content
        self._exported_func_names = exported_func_names
        self._exception = exception

    def get_fname(self):
        return self._fname

    def get_name(self):
        return TargetScript.to_name(self._fname)

    def get_content(self):
        return self._content

    def get_exported_func_names(self):
        return self._exported_func_names

    def get_exception(self):
        return self._exception


class ScriptsProcessor:
    """A script processor processes some file from a source dir and writes
    the target scripts to a target dir"""
    def __init__(self, logger, src_dir, target_dir, python_version):
        self._logger = logger
        self._src_dir = src_dir
        self._target_dir = target_dir
        self._python_version = python_version

    def process(self, script_fnames):
        """Explore the scripts. Since a script may import another script, we
        have a tree structure. We traverse the tree with classical DFS"""
        self._logger.log(logging.DEBUG, "Scripts to process: %s", script_fnames)

        self._scripts = []
        self._cur_script_fnames = list(script_fnames) # our stack.
        self._visited = set() # avoid cycles

        while self._has_more_scripts():
            self._process_next_script_if_not_visited()

        self._raise_exceptions()
        return self._scripts

    def _raise_exceptions(self):
        es = [script.get_exception() for script in self._scripts if script.get_exception()]
        if not es:
            return

        for e in es:
            self._logger.log(logging.ERROR, str(e))
        raise Exception("Compilation errors: see above")

    def append_script(self, script_fname):
        """Append a new script. The directive UseLib will call this method"""
        self._cur_script_fnames.append(script_fname)

    def _has_more_scripts(self):
        return len(self._cur_script_fnames)

    def _process_next_script_if_not_visited(self):
        next_script_fname = self._cur_script_fnames.pop()
        if next_script_fname in self._visited:
            return  # avoid cycles !

        self._process_script(next_script_fname)
        self._visited.add(next_script_fname)

    def _process_script(self, script_fname):
        directive_processor = DirectiveProcessor.create(self._src_dir, self, self._python_version)
        script_processor = ScriptProcessor(self._logger, directive_processor, script_fname)
        script = script_processor.parse_script()
        self._add_script(script)

    def _add_script(self, script):
        self._write_script(script)
        self._scripts.append(script)

    def get_exported_func_names_by_script_name(self):
        """Return a dict: script name -> exported functions"""
        return {script.get_name(): script.get_exported_func_names() for script in self._scripts}

    def _write_script(self, script):
        script_name = script.get_name()
        self._ensure_target_dir_exists()
        target_filename = os.path.join(self._target_dir, script_name)
        self._logger.log(logging.DEBUG, "Writing script: %s (%s)", script_name, target_filename)
        with open(target_filename, 'wb') as f:
            f.write(script.get_content())

    def _ensure_target_dir_exists(self):
        if not os.path.exists(self._target_dir):
            os.makedirs(self._target_dir)


class ScriptProcessor:
    """A script processor"""

    def __init__(self, logger, directive_processor, script_fname):
        self._logger = logger
        self._directive_processor = directive_processor
        self._script_fname = script_fname

    def parse_script(self):
        self._logger.log(logging.DEBUG, "Parsing script: %s (%s)", TargetScript.to_name(self._script_fname),
                         self._script_fname)
        exception = self._get_exception()
        content, exported_func_names = _ScriptParser(self._logger, self._directive_processor, self._script_fname)\
            .parse()
        return TargetScript(self._script_fname, content.encode("utf-8"), exported_func_names, exception)

    def _get_exception(self):
        try:
            py_compile.compile(self._script_fname, doraise=True)
        except Exception as e:
            return e
        else:
            return None


class _ScriptParser:
    """A script parser"""

    _PATTERN = re.compile("^def\\s+([^_].*?)\\(.*\\):.*$")

    def __init__(self, logger, directive_processor, script_fname):
        self._logger = logger
        self._directive_processor = directive_processor
        self._script_fname = script_fname
        self._script = None

    def parse(self):
        if self._script:
            return self._script

        self._exported_func_names = []
        self._lines = ["# parsed by py4lo (https://github.com/jferard/py4lo)"]
        with open(self._script_fname, 'r', encoding="utf-8") as f:
            self._process_lines(f)

        self._directive_processor.end()
        self._add_exported_func_names()
        return "\n".join(self._lines), self._exported_func_names

    def _process_lines(self, f):
        line = None
        try:
            for line in f:
                self._process_line(line.rstrip())
        except Exception as e:
            self._logger.critical(self._script_fname+", line="+line)
            raise e

    def _process_line(self, line):
        if line and line[0] == '#':
            self._lines.extend(self._directive_processor.process_line(line))
        else:
            self._append_to_lines(line)

    def _append_to_lines(self, line):
        if self._directive_processor.ignore_lines():
            self._lines.append("### py4lo ignore: " + line)
        else:
            self._update_exported(line)
            self._lines.append(line)

    def _update_exported(self, line):
        m = _ScriptParser._PATTERN.match(line)
        if m:
            func_name = m.group(1)
            self._exported_func_names.append(func_name)

    def _add_exported_func_names(self):
        if self._exported_func_names:
            self._lines.extend([
                "",
                "",
                "g_exportedScripts = ({},)".format(", ".join(self._exported_func_names))
            ])
