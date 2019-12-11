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
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, List, Set, Mapping, AbstractSet, Optional

from callbacks import AddReadmeWith
from core.asset import SourceAsset, DestinationAsset
from core.script import TempScript, DestinationScript, SourceScript
from toml_helper import load_toml


@dataclass
class Sources:
    source_ods_file: Path
    inc_dir: Path
    lib_dir: Path
    src_dir: Path
    src_ignore: List[str]
    opt_dir: Path
    assets_dir: Path
    assets_ignore: List[str]
    test_dir: Path

    def get_src_paths(self) -> Set[Path]:
        return _get_paths(self.src_dir, self.src_ignore, "*.py")

    def get_assets(self) -> List[SourceAsset]:
        return [SourceAsset(p, self.assets_dir) for p in
                _get_paths(self.assets_dir, self.assets_ignore)]

    def get_src_scripts(self) -> List[SourceScript]:
        script_paths = self.get_src_paths()
        return [SourceScript(sp, self.src_dir) for sp in script_paths]


@dataclass
class Destinations:
    dest_ods_file: Path
    temp_dir: Path
    dest_dir: Path
    assets_dest_dir: Path

    def to_destination_scripts(self, temp_scripts: List[TempScript]) -> List[
        DestinationScript]:
        return [ts.to_destination(self.dest_dir) for ts in
                temp_scripts]

    def to_destination_assets(self, source_assets) -> List[DestinationAsset]:
        return [sa.to_dest(self.assets_dest_dir) for sa in source_assets]


class PropertiesProvider:
    def __init__(self, logger: logging.Logger, base_path, sources: Sources,
                 destinations: Destinations, tdata: Mapping[str, Any]):
        self._base_path = base_path
        self._logger = logger
        self._tdata = tdata
        self._sources = sources
        self._destinations = destinations

    def get_base_path(self) -> Path:
        return self._base_path

    def get_logger(self) -> logging.Logger:
        return self._logger

    def keys(self) -> AbstractSet[str]:
        return self._tdata.keys()

    def get(self, k: str) -> Any:
        return self._tdata[k]

    def get(self, k: str, default: Optional[Any] = None) -> Any:
        return self._tdata.get(k, default)

    def get_sources(self) -> Sources:
        return self._sources

    def get_destinations(self) -> Destinations:
        return self._destinations

    def get_src_paths(self) -> Set[Path]:
        return self._sources.get_src_paths()

    def get_assets_paths(self) -> Set[Path]:
        return self._sources.get_assets_paths()

    def get_readme_callback(self) -> Optional[AddReadmeWith]:
        add_readme = self.get("add_readme", False)
        if add_readme:
            readme_contact = self.get("readme_contact")
            add_readme_callback = AddReadmeWith(
                self.get_base_path().joinpath("inc"),
                readme_contact)
        else:
            add_readme_callback = None
        return add_readme_callback


def _get_paths(source_dir: Path, ignore: List[str], glob="*") -> Set[Path]:
    paths = set(source_dir.rglob(glob))
    for pattern in ignore:
        paths -= set(source_dir.rglob(pattern))
    return set(p for p in paths if p.is_file())


class PropertiesProviderFactory:
    def create(self, toml_filename="py4lo.toml") -> PropertiesProvider:
        base_path = Path(__file__).parent.parent.parent.resolve()
        kwargs = {'py4lo': base_path, 'project': os.getcwd()}
        self._tdata = load_toml(base_path.joinpath(
            "default-py4lo.toml"), Path(toml_filename), kwargs)
        logger = self.get_logger()
        sources = self._get_sources()
        destinations = self._get_destinations(sources.source_ods_file)
        return PropertiesProvider(logger, base_path, sources, destinations,
                                  self._tdata)

    def _get_sources(self):
        src: Mapping[str, Any] = self._tdata["src"]
        sources = Sources(
            Path(src["source_ods_file"].format()),
            Path(src["inc_dir"]),
            Path(src["lib_dir"]),
            Path(src["src_dir"]),
            src["src_ignore"],
            Path(src["opt_dir"]),
            Path(src["assets_dir"]),
            src["assets_ignore"],
            Path(src["test_dir"]))
        return sources

    def _get_destinations(self, source_ods_file: Path):
        dest: Mapping[str, str] = self._tdata["dest"]
        destinations = Destinations(
            self._get_dest_file(source_ods_file),
            Path(dest["temp_dir"]), Path(dest["dest_dir"]),
            Path(dest["assets_dest_dir"]))
        return destinations

    def _get_dest_file(self, source_ods_file: Path) -> Path:
        dest: Mapping[str, str] = self._tdata["dest"]
        if "dest_ods_file" in dest:
            dest_ods_file = Path(dest["dest_ods_file"])
            if "suffix" in dest:
                self.get_logger().debug(
                    "Property dest_name set to `%s`, ignore suffix `%s`".format(
                        dest_ods_file, dest["suffix"]))
        else:
            suffix = dest["suffix"]
            dest_ods_file = source_ods_file.parent.joinpath(
                source_ods_file.stem + "-" + suffix + source_ods_file.suffix)

        return dest_ods_file

    def get_logger(self) -> logging.Logger:
        logging.basicConfig(level=logging.DEBUG)
        logger = logging.getLogger("py4lo")
        logger.setLevel(self._tdata["log_level"])
        return logger
