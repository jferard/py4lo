#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
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

from dataclasses import dataclass
from pathlib import Path


@dataclass
class DestinationAsset:
    path: Path
    content: bytes


@dataclass
class SourceAsset:
    path: Path
    assets_dir: Path

    def to_dest(self, assets_dest_dir) -> DestinationAsset:
        dest_path = assets_dest_dir.joinpath(
            self.path.relative_to(self.assets_dir)
        )
        with self.path.open("rb") as source:
            return DestinationAsset(dest_path, source.read())
