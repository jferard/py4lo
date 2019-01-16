# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2018 J. Férard <https://github.com/jferard>

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
    ods_dest_name = _get_dest_name(tdata)
    return OdsUpdater.create(tdata).update(ods_source_name, ods_dest_name)

def _get_logger(tdata):
    logger = logging.getLogger("py4lo")
    logger.setLevel(tdata["log_level"])
    return logger

def _get_dest_name(tdata):
    ods_source_name = tdata["source_file"]
    if "dest_name" in tdata:
        if "suffix" in tdata:
            _get_logger(tdata).debug("Property dest_name set to {}; ignore suffix {}".format(tdata["dest_name"], tdata["suffix"]))
        ods_dest_name = tdata["dest_name"]
    else:
        suffix = tdata["default_suffix"]
        ods_dest_name = ods_source_name[0:-4]+"-"+suffix+ods_source_name[-4:]
    return ods_dest_name

class Asset:
    def __init__(self, fname, content):
        self.__fname = fname
        self.__content = content

    def get_fname(self):
        return self.__fname

    def get_content(self):
        return self.__content

class OdsUpdater:
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
        assets_dir = tdata["assets_dir"]
        assets_dest_dir = tdata["assets_dest_dir"]
        target_dir = tdata["target_dir"]
        python_version = tdata["python_version"]

        return OdsUpdater(logger, src_dir, assets_dir, target_dir, assets_dest_dir, python_version, add_readme_callback)

    def __init__(self, logger, src_dir, assets_dir, target_dir, assets_dest_dir, python_version, add_readme_callback):
        self.__logger = logger
        self.__src_dir = src_dir
        self.__assets_dir = assets_dir
        self.__assets_dest_dir = assets_dest_dir
        self.__target_dir = target_dir
        self.__python_version = python_version
        self.__add_readme_callback = add_readme_callback

    def update(self, ods_source_name, ods_dest_name):
        self.__logger.info("Debug or init. Generating %s for Python %s", ods_dest_name, self.__python_version)

        scripts = self._get_scripts()
        assets = self._get_assets()

        zip_updater = self.__create_updater(scripts, assets)
        zip_updater.update(ods_source_name, ods_dest_name)
        return ods_dest_name

    def _get_scripts(self):
        script_fnames = set(
            os.path.join(self.__src_dir, fname) for fname in os.listdir(self.__src_dir) if fname.endswith(".py"))
        scripts_processor = ScriptsProcessor(self.__logger, self.__src_dir, self.__target_dir, self.__python_version)
        return scripts_processor.process(script_fnames)

    def _get_assets(self):
        assets = []
        for root, _, fnames in os.walk(self.__assets_dir):
            for fname in fnames:
                filename = os.path.join(root, fname)
                dest_name = os.path.join(self.__assets_dest_dir, os.path.relpath(filename, self.__assets_dir)).replace(
                    os.path.sep, "/")
                with open(filename, 'rb') as source:
                    assets.append(Asset(dest_name, source.read()))

        return assets

    def __create_updater(self, scripts, assets):
        zip_updater = ZipUpdater()
        (
            zip_updater
                .item(IgnoreScripts(ARC_SCRIPTS_PATH))
                .item(RewriteManifest(scripts, assets))
                .after(AddScripts(scripts))
                .after(AddAssets(assets))
        )
        if self.__add_readme_callback is not None:
            zip_updater.after(self.__add_readme_callback)

        return zip_updater

def open_with_calc(ods_name, calc_exe):
    """Open a file with calc"""
    _r = subprocess.call([calc_exe, ods_name])
