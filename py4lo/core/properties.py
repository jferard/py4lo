# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2024 J. Férard <https://github.com/jferard>
#
#     This file is part of Py4LO.
#
#     Py4LO is free software: you can redistribute it and/or modify
#     it under the terms of the GNU General Public License as published by
#     the Free Software Foundation, either version 3 of the License, or
#     (at your option) any later version.
#
#     Py4LO is distributed in the hope that it will be useful,
#     but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#     GNU General Public License for more details.
#
#     You should have received a copy of the GNU General Public License
#     along with this program.  If not, see <http://www.gnu.org/licenses/>.
import logging
import os
from pathlib import Path
from typing import Any, Set, Mapping, AbstractSet, Optional

from callbacks import AddReadmeWith
from core.source_dest import Sources, Destinations
from toml_helper import load_toml


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
        if "source_ods_file" in src:
            source_ods_file = Path(src["source_ods_file"])
        else:
            source_ods_file = None
        sources = Sources(
            source_ods_file,
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

    def _get_dest_file(self, source_ods_file: Optional[Path]) -> Path:
        dest: Mapping[str, str] = self._tdata["dest"]
        if "dest_ods_file" in dest:
            dest_ods_file = Path(dest["dest_ods_file"])
            if "suffix" in dest:
                self.get_logger().debug(
                    "Property dest_name set to `%s`, ignore suffix `%s`",
                    dest_ods_file, dest["suffix"])
        else:
            suffix = dest["suffix"]
            if source_ods_file is None:
                dest_ods_file = Path("new-project.ods")
            else:
                name = (source_ods_file.stem + "-"
                        + suffix + source_ods_file.suffix)
                dest_ods_file = source_ods_file.parent.joinpath(name)

        return dest_ods_file

    def get_logger(self) -> logging.Logger:
        logging.basicConfig(level=logging.DEBUG)
        logger = logging.getLogger("py4lo")
        logger.setLevel(self._tdata["log_level"])
        return logger
