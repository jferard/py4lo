# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2017 J. Férard <https://github.com/jferard>

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
from zip_updater import ZipUpdater
import logging
from callbacks import *
from scripts_processor import ScriptsProcessor
import subprocess

py4lo_path = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))

def update_ods(tdata):
    ods_source_name = tdata["source_file"]
    suffix = tdata["default_suffix"]
    ods_dest_name = ods_source_name[0:-4]+"-"+suffix+ods_source_name[-4:]
    return OdsUpdater.create(tdata).update(ods_source_name, ods_dest_name)

class OdsUpdater():
    @staticmethod
    def create(tdata):
        logger = logging.getLogger("py4lo")
        logger.setLevel(tdata["log_level"])
        add_readme = tdata["add_readme"]
        if add_readme:
            readme_contact = tdata["readme_contact"]
            add_readme_callback = AddReadmeWith(os.path.join(py4lo_path, "inc"), readme_contact)
        else:
            add_readme_callback = None

        src_dir = tdata["src_dir"]
        target_dir = tdata["target_dir"]
        python_version = tdata["python_version"]

        return OdsUpdater(logger, src_dir, target_dir, python_version, add_readme_callback)

    def __init__(self, logger, src_dir, target_dir, python_version, add_readme_callback):
        self.__logger = logger
        self.__src_dir = src_dir
        self.__target_dir = target_dir
        self.__python_version = python_version
        self.__add_readme_callback = add_readme_callback

    def update(self, ods_source_name, ods_dest_name):
        self.__logger.info("Debug or init. Generating %s for Python %s", ods_dest_name, self.__python_version)

        script_fnames = set(os.path.join(self.__src_dir, fname) for fname in os.listdir(self.__src_dir) if fname.endswith(".py"))
        scripts_processor = ScriptsProcessor(self.__logger, self.__src_dir, self.__python_version, self.__target_dir)
        scripts = scripts_processor.process(script_fnames)

        zip_updater = self.__create_updater(scripts)
        zip_updater.update(ods_source_name, ods_dest_name)
        return ods_dest_name

    def __create_updater(self, scripts):
        zip_updater = ZipUpdater()
        (
            zip_updater
                .item(IgnoreScripts(ARC_SCRIPTS_PATH))
                .item(RewriteManifest(scripts))
                .after(AddScripts(scripts))
        )
        if self.__add_readme_callback is not None:
            zip_updater.after(self.__add_readme_callback)

        return zip_updater

def open_with_calc(ods_name, calc_exe):
    r = subprocess.call([calc_exe, ods_name])
