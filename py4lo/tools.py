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
import logging
from callbacks import *
from script_processor import ScriptProcessor
import subprocess

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

def update_ods(tdata):
    ods_source_name = tdata["source_file"]
    suffix = tdata["default_suffix"]
    add_readme = tdata["add_readme"]
    readme_contact = tdata["readme_contact"]
    src_dir = tdata["src_dir"]
    target_dir = tdata["target_dir"]
    python_version = tdata["python_version"]

    logger = logging.getLogger("py4lo")
    logging.info("")
    logger.setLevel(tdata["log_level"])
    ods_dest_name = ods_source_name[0:-4]+"-"+suffix+ods_source_name[-4:]

    logger.setLevel(tdata["log_level"])
    logger.setLevel("DEBUG")
    logger.info("Debug or init. Generating %s for Python %s", ods_dest_name, python_version)
    
    script_fnames = set(os.path.join(src_dir, fname) for fname in os.listdir(src_dir) if fname.endswith(".py"))
    script_processor = ScriptProcessor(logger, src_dir, python_version, target_dir)
    script_processor.process(script_fnames)

    item_cbs = [ignore_scripts, rewrite_manifest(script_processor.get_scripts())]
    after_cbs = [add_scripts(script_processor.get_scripts())]
    if add_readme:
        after_cbs.append(add_readme_cb(readme_contact))
        
    update_zip(ods_source_name, ods_dest_name, item_callbacks = item_cbs, after_callbacks = after_cbs)
    return ods_dest_name

def debug_scripts(tdata, file_key):
    # retrieve infos from data
    target_dir = tdata["target_dir"]
    debug_path = os.path.join(target_dir, tdata[file_key])
    src_dir = tdata["src_dir"]
    python_version = tdata["python_version"]
    ods_dest_name = tdata[file_key]
    
    logger = logging.getLogger("py4lo")
    logging.info("")
    logger.setLevel(tdata["log_level"])
    logger.info("Debug or init. Generating %s for Python %s", debug_path, python_version)
    
    script_fnames = set(os.path.join(src_dir, fname) for fname in os.listdir(src_dir) if fname.endswith(".py"))
    script_processor = ScriptProcessor(logger, src_dir, python_version, target_dir)
    script_processor.process(script_fnames)
    
    item_cbs = [ignore_scripts, rewrite_manifest(script_processor.get_scripts())]
    after_cbs = [add_scripts(script_processor.get_scripts()), add_debug_content(script_processor.get_exported_func_names_by_script_name())]
    update_zip(os.path.join(tdata["py4lo_path"], "inc", "debug.ods"), ods_dest_name, item_callbacks = item_cbs, after_callbacks = after_cbs)
    return ods_dest_name
    
def open_with_calc(ods_name, calc_exe):
    r = subprocess.call([calc_exe, ods_name])
def test_all(tdata):
    final_status = 0
    for dirpath, dirnames, filenames in os.walk(tdata["test_dir"]):
        for filename in filenames:
            if filename.endswith("_test.py"):
                cmd = "\""+tdata["python_exe"]+"\" "+os.path.join(dirpath, filename)
                print (cmd)
                status, output = subprocess.getstatusoutput("\""+tdata["python_exe"]+"\" "+os.path.join(dirpath, filename))
                print (output)
                if status != 0:
                    final_status = 1

    return final_status
                    