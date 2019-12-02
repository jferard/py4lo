# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2019 J. Férard <https://github.com/jferard>

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
import logging
import subprocess
from pathlib import Path
from typing import List, Collection

from asset import Asset
from callbacks import *
from zip_updater import ZipUpdater
from script_set_processor import ScriptSetProcessor

py4lo_path = Path(__file__).parent


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
            _get_logger(tdata).debug(
                "Property dest_name set to {}; ignore suffix {}".format(
                    tdata["dest_name"], tdata["suffix"]))
        ods_dest_name = tdata["dest_name"]
    else:
        suffix = tdata["default_suffix"]
        ods_dest_name = ods_source_name[0:-4] + "-" + suffix + ods_source_name[
                                                               -4:]
    return ods_dest_name


def get_assets(assets_dir: Path, assets_ignore: List[str],
               assets_dest_dir: Path) -> List[Asset]:
    assets = []
    for p in get_paths(assets_dir, assets_ignore):
        dest = assets_dest_dir.joinpath(p.relative_to(assets_dir))
        with dest.open('rb') as source:
            assets.append(Asset(dest, source.read()))

    return assets


def get_paths(source_dir: Path, ignore: List[str], glob="*") -> Collection[
    Path]:
    paths = set(source_dir.rglob(glob))
    for pattern in ignore:
        paths -= set(source_dir.rglob(pattern))
    return set(p for p in paths if p.is_file())


class OdsUpdater:
    @staticmethod
    def create(tdata):
        logger = logging.getLogger("py4lo")
        logger.setLevel(tdata["log_level"])
        add_readme = tdata["add_readme"]
        if add_readme:
            readme_contact = tdata["readme_contact"]
            add_readme_callback = AddReadmeWith(py4lo_path.joinpath("inc"),
                                                readme_contact)
        else:
            add_readme_callback = None

        src_dir = tdata["src_dir"]
        src_ignore = tdata["src_ignore"]
        assets_dir = tdata["assets_dir"]
        assets_ignore = tdata["assets_ignore"]
        assets_dest_dir = tdata["assets_dest_dir"]
        target_dir = tdata["target_dir"]
        python_version = tdata["python_version"]

        return OdsUpdater(logger, Path(src_dir), src_ignore, Path(assets_dir),
                          assets_ignore, Path(target_dir),
                          Path(assets_dest_dir), python_version,
                          add_readme_callback)

    def __init__(self, logger: logging.Logger, src_dir: Path,
                 src_ignore: List[str], assets_dir: Path,
                 assets_ignore: List[str],
                 target_dir: Path, assets_dest_dir: Path,
                 python_version: str, add_readme_callback: AddReadmeWith):
        self._logger = logger
        self._src_dir = src_dir
        self._src_ignore = src_ignore
        self._assets_dir = assets_dir
        self._assets_ignore = assets_ignore
        self._assets_dest_dir = assets_dest_dir
        self._target_dir = target_dir
        self._python_version = python_version
        self._add_readme_callback = add_readme_callback

    def update(self, ods_source: Path, ods_dest: Path) -> Path:
        self._logger.info("Debug or init. Generating %s for Python %s",
                          ods_dest, self._python_version)

        scripts = self._get_scripts()
        assets = get_assets(self._assets_dir, self._assets_ignore,
                            self._assets_dest_dir)

        zip_updater = self._create_updater(scripts, assets)
        zip_updater.update(ods_source, ods_dest)
        return ods_dest

    def _get_scripts(self):
        script_paths = get_paths(self._src_dir, self._src_ignore, "*.py")
        scripts_processor = ScriptSetProcessor(self._logger, self._src_dir,
                                               self._target_dir,
                                               self._python_version,
                                               script_paths)
        return scripts_processor.process()

    def _create_updater(self, scripts, assets):
        zip_updater = ZipUpdater()
        (
            zip_updater.item(IgnoreScripts(ARC_SCRIPTS_PATH))
                .item(RewriteManifest(scripts, assets))
                .after(AddScripts(scripts))
                .after(AddAssets(assets))
        )
        if self._add_readme_callback is not None:
            zip_updater.after(self._add_readme_callback)

        return zip_updater


def open_with_calc(ods_name: Path, calc_exe):
    """Open a file with calc"""
    _r = subprocess.call([calc_exe, str(ods_name)])
