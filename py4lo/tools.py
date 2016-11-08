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
import zipfile
from callbacks import *
from script_processor import ScriptProcessor

def update_zip(zip_source_name,zip_dest_name, before_callbacks = [],
	item_callbacks = [], after_callbacks = [] ):
	with zipfile.ZipFile(zip_dest_name, 'w') as zout:
		for before_callback in before_callbacks:
			if not before_callback(zout):
				break
		with zipfile.ZipFile(zip_source_name, 'r') as zin:
			zout.comment = zin.comment # preserve the comment
			for item in zin.infolist():
				for item_callback in item_callbacks:
					if not item_callback(zin, zout, item):
						break
		for after_callback in after_callbacks:
			if not after_callback(zout):
				break

def update_ods(ods_source_name, suffix = "generated", readme = False, contact = "<add a contact in default-py4lo.toml file>", scripts_path=".", python_version="3"):
	ods_dest_name = ods_source_name[0:-4]+"-"+suffix+ods_source_name[-4:]
	script_fnames = set(os.path.join(scripts_path, f) for f in os.listdir(scripts_path) if f.endswith(".py"))
	script_processor = ScriptProcessor(scripts_path, python_version)
	script_processor.process(script_fnames)

	item_cbs = [ignore_scripts, rewrite_manifest(script_processor.get_scripts())]
	after_cbs = [add_scripts(script_processor.get_scripts())]
	if readme:
		after_cbs.append(add_readme(contact))
	update_zip(ods_source_name, ods_dest_name, item_callbacks = item_cbs, after_callbacks = after_cbs)
	return ods_dest_name

def debug_scripts(debug_file, scripts_path=".", python_version="3"):
	ods_dest_name = debug_file
	script_fnames = set(os.path.join(scripts_path, f) for f in os.listdir(scripts_path) if f.endswith(".py"))
	script_processor = ScriptProcessor(scripts_path, python_version)
	script_processor.process(script_fnames)
	
	py4lo_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "..")

	item_cbs = [ignore_scripts, rewrite_manifest(script_processor.get_scripts())]
	after_cbs = [add_scripts(script_processor.get_scripts()), add_debug_content(script_processor.get_func_names_by_script_name())]
	update_zip(os.path.join(py4lo_path, "inc", "debug.ods"), ods_dest_name, item_callbacks = item_cbs, after_callbacks = after_cbs)
	return ods_dest_name
	
def open_with_calc(ods_name, calc_exe):
	import subprocess
	r = subprocess.call([calc_exe, ods_name])