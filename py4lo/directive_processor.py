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

class DirectiveProcessor():
	def __init__(self, python_version, scripts_path, cur_script_fnames):
		self.__python_version = python_version
		self.__scripts_path = scripts_path
		self.__cur_script_fnames = cur_script_fnames
		self.__py4lo_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "..")
		self.__bootstrap = self.__get_bootstrap()
		
	def new_script(self):
		self.__bootstrapped = False

	def process_line(self, line):
		s = ""
		try:
			ls = shlex.split(line)
			if len(ls) >= 2 and ls[0] == '#' and ls[1] == 'py4lo:':
				directive = ls[2]
				args = ls[3:]
				if directive == 'use':
					if not self.__bootstrapped:
						s += self.__bootstrap
						self.__bootstrapped = True

					if args[0] == 'lib':
						s += self.__use_lib(args[1:])
					else:
						s += self.__use_object(args)
			else:
				s += line
		except ValueError:
			print ("non parsable line " + line)
			s += line
		
		return s

	def __use_lib(self, args):
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
