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
import shlex
import re

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
		self.__py4lo_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "..")
		self.__bootstrap = None

	def process(self, script_fnames):
		self.__scripts = []
		self.__func_names_by_script_name = {}
		self.__cur_script_fnames = [(fname, os.path.split(fname)[1]) for fname in script_fnames]
		visited = set()
		while len(self.__cur_script_fnames):
			(script_fname, script_name) = self.__cur_script_fnames.pop()
			if script_fname in visited:
				continue

			self.__parse_script(script_fname, script_name)
			visited.add(script_fname)

		return self.__scripts

	def __parse_script(self, script_fname, script_name):
		func_names = []

		s = "# parsed by py4lo\n"
		bootstrapped = False
		with open(script_fname, 'r', encoding="utf-8") as f:
			for line in f:
				m = re.match("^def\s+(.*?)\(.*\):\s*$", line)
				if m:
					func_name = m.group(1)
					if func_name[0] != "_":
						func_names.append(func_name)
					s += line
				elif line[0] == '#':
					try:
						ls = shlex.split(line)
						if len(ls) >= 2 and ls[0] == '#' and ls[1] == 'py4lo:':
							directive = ls[2]
							args = ls[3:]
							if directive == 'use':
								if not bootstrapped:
									if self.__bootstrap is None:
										self.__bootstrap = self.__get_bootstrap()
									s += self.__bootstrap
									bootstrapped = True

								if args[0] == 'lib':
									s += self.__use_lib(args[1:])
								else:
									s += self.__use_object(args)
						else:
							s += line
					except ValueError:
						print ("non parsable line " + line)
						s += line
				else:
					s += line

		s += "\n\ng_exportedScripts = ("+", ".join(func_names)+")\n"
		self.__func_names_by_script_name[script_name] = func_names

		script = Script(script_fname, script_name, s.encode("utf-8"))
		self.__scripts.append(script)

	def __use_lib(self, args):
		object_ref = args[0]
		(lib_ref, object_name) = object_ref.split("::")
		script_fname_wo_extension = os.path.join(self.__py4lo_path, "lib", lib_ref)

		script_fname = script_fname_wo_extension + "__" + self.__python_version + ".py"
		if not os.path.isfile(script_fname):
			script_fname = script_fname_wo_extension + ".py"
			if not os.path.isfile(script_fname):
				raise BaseException(script_fname + " is not a py4lo lib")

		if len(args) == 3 and args[1] == 'as':
			ret = args[2]
		else:
			ret = object_name

		self.__cur_script_fnames.append((script_fname, lib_ref+".py"))
		return ret + " = use_local(\""+object_ref+"\")\n"

	def __use_object(self, args):
		object_ref = args[0]
		(script_ref, object_name) = object_ref.split("::")
		script_fname = os.path.join(self.__scripts_path, "", script_ref+".py")
		if not os.path.isfile(script_fname):
			raise BaseException(script_fname + " is not a script")

		if len(args) == 3 and args[1] == 'as':
			ret = args[2]
		else:
			ret = object_name

		self.__cur_script_fnames.append((script_fname, script_ref+".py"))
		return ret + " = use_local(\""+object_ref+"\")\n"

	def __get_bootstrap(self):
		path = os.path.dirname(os.path.realpath(__file__))

		bootstrap = ""
		state = 0
		with open(os.path.join(path, "..", "inc", "py4lo_bootstrap.py"), 'r', encoding="utf-8") as b:
			for line in b.readlines():
				l = line.strip()
				if state == 0:
					if l.startswith('#'):
						pass
					elif l.startswith("\"\"\""):
						state = 1
					else:
						bootstrap += line
				elif state == 1:
					if l.endswith("\"\"\""):
						state = 0
		return bootstrap + "\n"

	def get_scripts(self):
		return self.__scripts
		
	def get_func_names_by_script_name(self):
		return self.__func_names_by_script_name