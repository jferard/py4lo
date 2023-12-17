# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
#
#     This file is part of Py4LO.
#
#     Py4LO is free software: you can redistribute it and/or modify
#     it under the terms of the GNU General Public License as published by
#     the Free Software Foundation, either version 3 of the License, or
#     (at your option) any later version.
#
#     Py4LO is distributed in the hope that it will be useful,
#     but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#     GNU General Public License for more details.
#
#     You should have received a copy of the GNU General Public License
#     along with this program.  If not, see <http://www.gnu.org/licenses/>.
import logging
import py_compile
import re
from pathlib import Path
from typing import List, Optional, Sequence, cast, Set

from core.script import TempScript, ParsedScriptContent, SourceScript
from directive_processor import DirectiveProcessor
from directives import DirectiveProvider


class ScriptSetProcessor:
    """
    A script processor processes some file from a source dir and writes
    the target scripts to a target dir
    """

    def __init__(
        self,
        logger: logging.Logger,
        target_dir: Path,
        python_version: str,
        directive_provider: DirectiveProvider,
        source_scripts: Sequence[SourceScript],
    ):
        self._logger = logger
        self._target_dir = target_dir
        self._python_version = python_version
        self._directive_provider = directive_provider
        self._source_scripts = source_scripts
        self._logger.debug("Scripts to process: %s", source_scripts)
        self._scripts: List[TempScript] = []
        self._cur_source_scripts = list(source_scripts)  # our stack.
        self._visited = cast(Set[SourceScript], set())  # avoid cycles

    def process(self) -> List[TempScript]:
        """Explore the scripts. Since a script may import another script, we
        have a tree structure. We traverse the tree with classical DFS"""
        while self._has_more_scripts():
            self._process_next_script_if_not_visited()

        self._raise_exceptions()
        return self._scripts

    def _has_more_scripts(self) -> bool:
        return bool(self._cur_source_scripts)

    def _process_next_script_if_not_visited(self):
        next_script = self._cur_source_scripts.pop()
        if next_script in self._visited:
            return  # avoid cycles !

        self._process_script(next_script)
        self._visited.add(next_script)

    def _process_script(self, source_script: SourceScript):
        directive_processor = DirectiveProcessor.create(
            self, self._directive_provider, self._python_version, source_script
        )
        script_processor = ScriptProcessor(
            self._logger, directive_processor, source_script, self._target_dir
        )
        temp_script = script_processor.parse_script()
        self.add_script(temp_script)

    def _raise_exceptions(self):
        es = [script.exception for script in self._scripts if script.exception]
        if not es:
            return

        for e in es:
            self._logger.error(str(e))
        raise Exception("Compilation errors: see above")

    def append_script(self, source_script: SourceScript):
        """Append a new script. The directive UseLib will call this method"""
        self._cur_source_scripts.append(source_script)

    def add_script(self, script: TempScript):
        """Add but do not process this script"""
        self._write_script(script)
        self._scripts.append(script)

    def _write_script(self, script: TempScript):
        self._ensure_dir_exists(script)
        self._logger.debug(
            "Writing temp script: %s (%s)", script.relative_path, script.script_path
        )
        with script.script_path.open("wb") as f:
            f.write(script.script_content)

    def _ensure_dir_exists(self, script: TempScript):
        target_dir = script.script_path.parent
        if not target_dir.exists():
            target_dir.mkdir(parents=True)


class ScriptProcessor:
    """A script processor"""

    def __init__(
        self,
        logger: logging.Logger,
        directive_processor: DirectiveProcessor,
        source_script: SourceScript,
        target_dir: Path,
    ):
        self._logger = logger
        self._directive_processor = directive_processor
        self._source_script = source_script
        self._target_dir = target_dir

    def parse_script(self) -> TempScript:
        self._logger.debug(
            "Parsing script: %s (%s)",
            self._source_script.relative_path,
            self._source_script,
        )
        exception = self._get_exception()
        parser = _ContentParser(
            self._logger, self._directive_processor, self._source_script.script_path
        )
        parsed_content = parser.parse(self._source_script.export_funcs)
        target_path = self._target_dir.joinpath(self._source_script.relative_path)
        script = TempScript(
            target_path,
            parsed_content.text.encode("utf-8"),
            self._target_dir,
            parsed_content.exported_func_names,
            exception,
        )
        self._logger.debug(
            "Temp output script is: %s (%s)",
            script.script_path,
            script.exported_func_names,
        )
        return script

    def _get_exception(self) -> Optional[Exception]:
        try:
            py_compile.compile(str(self._source_script.script_path), doraise=True)
        except Exception as e:
            return e
        else:
            return None


class _ContentParser:
    """A script parser"""

    _PATTERN = re.compile("^def\\s+([^_].*?)\\(.*\\):.*$")

    def __init__(
        self,
        logger: logging.Logger,
        directive_processor: DirectiveProcessor,
        script_path: Path,
    ):
        self._logger = logger
        self._directive_processor = directive_processor
        self._script_path = script_path
        self._script = None
        self._exported_func_names = cast(List[str], [])
        self._lines = ["# parsed by py4lo (https://github.com/jferard/py4lo)"]

    def parse(self, export_funcs: bool) -> ParsedScriptContent:
        if self._script:
            return self._script

        with self._script_path.open("r", encoding="utf-8") as f:
            self._process_lines(f)

        self._directive_processor.end()
        if export_funcs:
            exported_func_names = self._exported_func_names
            self._add_exported_func_names()
        else:
            exported_func_names = []
        return ParsedScriptContent("\n".join(self._lines), exported_func_names)

    def _process_lines(self, f):
        line = None
        try:
            for line in f:
                self._process_line(line.rstrip())
        except Exception as e:
            self._logger.critical("%s, line=%s", self._script_path, line)
            raise e

    def _process_line(self, line: str):
        if line and line[0] == "#":
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
        m = _ContentParser._PATTERN.match(line)
        if m:
            func_name = m.group(1)
            self._exported_func_names.append(func_name)

    def _add_exported_func_names(self):
        if self._exported_func_names:
            self._lines.extend(
                [
                    "",
                    "",
                    "g_exportedScripts = ({},)".format(
                        ", ".join(self._exported_func_names)
                    ),
                ]
            )
