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
import os
import re
from directive_processor import DirectiveProcessor

class Script():
	def __init__(self, script_fname, script_name, script_data):
		self.__fname = script_fname
		self.__name = script_name
		self.__data = script_data

	def get_fname(self):
		return self.__fname

	def get_name(self):
		return self.__name

	def get_data(self):
		return self.__data

class ScriptProcessor():
	def __init__(self, scripts_path, python_version):
		self.__scripts_path = scripts_path
		self.__python_version = python_version

	def process(self, script_fnames):
		self.__scripts = []
		self.__func_names_by_script_name = {}
		cur_script_fnames = [(fname, os.path.split(fname)[1]) for fname in script_fnames]
		self.__directive_processor = DirectiveProcessor(self.__python_version, self.__scripts_path, cur_script_fnames)
		visited = set()
		while len(cur_script_fnames):
			(script_fname, script_name) = cur_script_fnames.pop()
			if script_fname in visited:
				continue

			self.__parse_script(script_fname, script_name)
			visited.add(script_fname)

		return self.__scripts

	def __parse_script(self, script_fname, script_name):
		func_names = []

		s = "# parsed by py4lo\n"
		self.__directive_processor.new_script()
		with open(script_fname, 'r', encoding="utf-8") as f:
			for line in f:
				m = re.match("^def\s+(.*?)\(.*\):\s*$", line)
				if m:
					func_name = m.group(1)
					if func_name[0] != "_":
						func_names.append(func_name)
					s += line
				elif line[0] == '#':
					s += self.__directive_processor.process_line(line)
				else:
					s += line

		s += "\n\ng_exportedScripts = ("+", ".join(func_names)+")\n"
		self.__func_names_by_script_name[script_name] = func_names

		script = Script(script_fname, script_name, s.encode("utf-8"))
		self.__scripts.append(script)
		
	def get_scripts(self):
		return self.__scripts
		
	def get_func_names_by_script_name(self):
		return self.__func_names_by_script_name