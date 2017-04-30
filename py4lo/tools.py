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
from zip_updater import ZipUpdater
import logging
from callbacks import *
from script_processor import ScriptProcessor
import subprocess 

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

    scripts = script_processor.get_scripts()

    zip_updater = ZipUpdater()
    (
        zip_updater
            .item(ignore_scripts)
            .item(rewrite_manifest(scripts))
            .after(add_scripts(scripts))
    )
    if add_readme:
        zip_updater.after(add_readme_cb(readme_contact))

    zip_updater.update(ods_source_name, ods_dest_name)
    return ods_dest_name

def open_with_calc(ods_name, calc_exe):
    r = subprocess.call([calc_exe, ods_name])
