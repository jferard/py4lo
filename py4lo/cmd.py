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
import toml
import os
import sys
from tools import update_ods, open_with_calc, debug_scripts

def load_toml(fname = "py4lo.toml"):
	py4lo_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "..")
	with open(os.path.join(py4lo_path, "default-py4lo.toml"), 'r', encoding="utf-8") as default_py4lo_toml:
		data = toml.loads(default_py4lo_toml.read())
	
	try:
		with open(fname, 'r', encoding="utf-8") as py4lo_toml:
			temp = toml.loads(py4lo_toml.read())
	except OSError as ose:
		pass
	else:
		data.update(temp)
	
	if not "python_version" in data:
		data["python_version"] = str(sys.version_info.major)

	return data
  
def debug():
	tdata = load_toml()
	return debug_scripts(tdata["debug_file"], ".", tdata["python_version"])

def init():
	tdata = load_toml()
	return debug_scripts(tdata["init_file"], ".", tdata["python_version"])
	
def test():
	tdata = load_toml()
	dest_name = update()
	open_with_calc(dest_name, tdata["calc_exe"])

def update():
	tdata = load_toml()
	return update_ods(tdata["source_file"], tdata["default_suffix"], tdata["add_readme"], tdata["readme_contact"], ".", tdata["python_version"])