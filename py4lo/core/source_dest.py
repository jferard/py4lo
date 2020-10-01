#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2019 J. Férard <https://github.com/jferard>
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
import os
from dataclasses import dataclass
from pathlib import Path
from typing import List, Set

from core.asset import SourceAsset, DestinationAsset
from core.script import SourceScript, TempScript, DestinationScript


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

    def get_module_names(self) -> Set[str]:
        return _get_module_names(self.src_dir, self.src_ignore, "*.py")

    def get_all_module_names(self) -> Set[str]:
        return {*_get_module_names(self.src_dir, self.src_ignore, "*.py"),
                *_get_module_names(self.lib_dir, self.src_ignore),
                *_get_module_names(self.opt_dir, self.src_ignore)}

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


def _get_paths(source_dir: Path, ignore: List[str], glob="*") -> Set[Path]:
    paths = set(source_dir.rglob(glob))
    for pattern in ignore:
        paths -= set(source_dir.rglob(pattern))
    return set(p for p in paths if p.is_file())


def _get_module_names(source_dir: Path, ignore: List[str], glob="*"
                      ) -> Set[str]:
    paths = _get_paths(source_dir, ignore, glob)
    module_names = set()
    for p in paths:
        p = p.relative_to(source_dir).with_suffix("")
        if p.parts[0] == "__pycache__":
            continue

        if p.name in ("__main__", "__init__"):
            p = p.parent

        module_names.add(str(p).replace(os.path.sep, "."))
    return module_names
