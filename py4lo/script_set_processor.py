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
import logging
import py_compile
import re
from dataclasses import dataclass
from pathlib import Path
from typing import List, Collection, Optional, Dict

from directive_processor import DirectiveProcessor


@dataclass
class TargetScript:
    """A target script has a file name, a content (the file was processed),
    some exported functions and the exceptions encountered by the interpreter"""
    script_path: Path
    script_content: bytes
    exported_func_names: List[str]
    exception: Optional[Exception]

    @property
    def name(self):
        return self.script_path.stem


class ScriptSetProcessor:
    """A script processor processes some file from a source dir and writes
    the target scripts to a target dir"""

    def __init__(self, logger: logging.Logger, src_dir: Path, target_dir: Path,
                 python_version: str,
                 script_paths: Collection[Path]):
        self._logger = logger
        self._src_dir = src_dir
        self._target_dir = target_dir
        self._python_version = python_version
        self._logger.log(logging.DEBUG, "Scripts to process: %s", script_paths)
        self._scripts: List[TargetScript] = []
        self._cur_script_paths = list(script_paths)  # our stack.
        self._visited = set()  # avoid cycles

    def process(self) -> List[TargetScript]:
        """Explore the scripts. Since a script may import another script, we
        have a tree structure. We traverse the tree with classical DFS"""
        while self._has_more_scripts():
            self._process_next_script_if_not_visited()

        self._raise_exceptions()
        return self._scripts

    def _has_more_scripts(self) -> bool:
        return bool(self._cur_script_paths)

    def _process_next_script_if_not_visited(self):
        next_script_path = self._cur_script_paths.pop()
        if next_script_path in self._visited:
            return  # avoid cycles !

        self._process_script(next_script_path)
        self._visited.add(next_script_path)

    def _process_script(self, script_path: Path):
        directive_processor = DirectiveProcessor.create(self._src_dir, self,
                                                        self._python_version)
        script_processor = ScriptProcessor(self._logger, directive_processor,
                                           script_path, self._target_dir)
        script = script_processor.parse_script()
        self._add_script(script)

    def _raise_exceptions(self):
        es = [script.exception for script in self._scripts if
              script.exception]
        if not es:
            return

        for e in es:
            self._logger.log(logging.ERROR, str(e))
        raise Exception("Compilation errors: see above")

    def append_script(self, script_path: Path):
        """Append a new script. The directive UseLib will call this method"""
        self._cur_script_paths.append(script_path)

    def _add_script(self, script: TargetScript):
        self._write_script(script)
        self._scripts.append(script)

    def get_exported_func_names_by_script_name(self) -> Dict[str, List[str]]:
        """Return a dict: script name -> exported functions"""
        return {script.name: script.exported_func_names for script
                in self._scripts}

    def _write_script(self, script: TargetScript):
        script_name = script.name
        self._ensure_target_dir_exists()
        self._logger.log(logging.DEBUG, "Writing script: %s (%s)", script_name,
                         script.script_path)
        with script.script_path.open('wb') as f:
            f.write(script.script_content)

    def _ensure_target_dir_exists(self):
        if not self._target_dir.exists():
            self._target_dir.mkdir(parents=True)


class ScriptProcessor:
    """A script processor"""

    def __init__(self, logger: logging.Logger,
                 directive_processor: DirectiveProcessor, script_path: Path,
                 target_dir: Path):
        self._logger = logger
        self._directive_processor = directive_processor
        self._script_path = script_path
        self._target_dir = target_dir

    def parse_script(self) -> TargetScript:
        self._logger.log(logging.DEBUG, "Parsing script: %s (%s)",
                         self._script_path.stem,
                         self._script_path)
        exception = self._get_exception()
        content, exported_func_names = _ScriptParser(self._logger,
                                                     self._directive_processor,
                                                     self._script_path) \
            .parse()
        target_path = self._target_dir.joinpath(self._script_path.name)
        return TargetScript(target_path, content.encode("utf-8"),
                            exported_func_names, exception)

    def _get_exception(self) -> Optional[Exception]:
        try:
            py_compile.compile(self._script_path, doraise=True)
        except Exception as e:
            return e
        else:
            return None


class _ScriptParser:
    """A script parser"""

    _PATTERN = re.compile("^def\\s+([^_].*?)\\(.*\\):.*$")

    def __init__(self, logger: logging.Logger,
                 directive_processor: DirectiveProcessor,
                 script_path: Path):
        self._logger = logger
        self._directive_processor = directive_processor
        self._script_path = script_path
        self._script = None
        self._exported_func_names = []
        self._lines = ["# parsed by py4lo (https://github.com/jferard/py4lo)"]

    def parse(self) -> (str, List[str]):
        if self._script:
            return self._script

        with self._script_path.open('r', encoding="utf-8") as f:
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
            self._logger.critical("%s, line=%s", self._script_path, line)
            raise e

    def _process_line(self, line: str):
        if line and line[0] == '#':
            self._lines.extend(self._directive_processor.process_line(line))
        else:
            self._append_to_lines(line)

    def _append_to_lines(self, line: str):
        if self._directive_processor.ignore_lines():
            self._lines.append("### py4lo ignore: " + line)
        else:
            self._update_exported(line)
            self._lines.append(line)

    def _update_exported(self, line: str):
        m = _ScriptParser._PATTERN.match(line)
        if m:
            func_name = m.group(1)
            self._exported_func_names.append(func_name)

    def _add_exported_func_names(self):
        if self._exported_func_names:
            self._lines.extend([
                "",
                "",
                "g_exportedScripts = ({},)".format(
                    ", ".join(self._exported_func_names))
            ])
