# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2021 J. Férard <https://github.com/jferard>

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
import xml.dom.minidom
from pathlib import Path
from typing import List
from zipfile import ZipFile, ZipInfo

from callbacks.callback import ItemCallback
from core.asset import DestinationAsset
from core.script import DestinationScript

BASIC_ENTRIES = [
    """<manifest:file-entry manifest:full-path="Basic" manifest:media-type="application/binary"/>""",
    """<manifest:file-entry manifest:full-path="Basic/Standard" manifest:media-type="application/binary"/>""",
    """<manifest:file-entry manifest:full-path="Basic/Standard/py4lo.xml" manifest:media-type="text/xml"/>""",
    """<manifest:file-entry manifest:full-path="Basic/Standard/script-lb.xml" manifest:media-type="text/xml"/>""",
    """<manifest:file-entry manifest:full-path="Basic/script-lc.xml" manifest:media-type="text/xml"/>""",
]

PYTHON_ENTRY_TPL = """<manifest:file-entry manifest:full-path="{0}" manifest:media-type=""/>"""
DIR_TPL = """<manifest:file-entry manifest:full-path="{0}" manifest:media-type="application/binary"/>"""
ASSET_ENTRY_TPL = """<manifest:file-entry manifest:full-path="{0}" manifest:media-type="application/octet-stream"/>"""
MANIFEST_CLOSE_TAG = """</manifest:manifest>"""


class RewriteManifest(ItemCallback):
    def __init__(self, scripts: List[DestinationScript],
                 assets: List[DestinationAsset]):
        self._scripts = scripts
        self._assets = assets

    def call(self, zin: ZipFile, zout: ZipFile, item: ZipInfo) -> bool:
        if item.filename == "META-INF/manifest.xml":
            data = self._rewrite_manifest(zin, item)
            zout.writestr(item.filename, data)
            return True
        else:
            return False

    def _rewrite_manifest(self, zin: ZipFile, manifest_item: ZipInfo):
        pretty_manifest = RewriteManifest._prettyfy_xml(zin,
                                                        manifest_item.filename)
        lines = RewriteManifest._strip_close(pretty_manifest)
        lines += self._script_lines()
        return "\n".join(lines).encode("utf-8")

    @staticmethod
    def _prettyfy_xml(zin: ZipFile, fname: str):
        data = zin.read(fname)
        xml_str = xml.dom.minidom.parseString(data.decode("utf-8"))
        return xml_str.toprettyxml(indent="   ", newl='')

    @staticmethod
    def _strip_close(pretty_manifest: str):
        lines = []
        for line in pretty_manifest.splitlines():
            if line.strip() == MANIFEST_CLOSE_TAG:  # end of manifest
                return lines
            lines.append(line)

        raise Exception("no manifest closing tag in " + pretty_manifest)

    def _script_lines(self):
        lines = ["    " + be for be in BASIC_ENTRIES]

        for d in self._get_dirs():
            lines.append("    " + DIR_TPL.format(d.as_posix()))
        for script in self._scripts:
            lines.append(
                "    " + PYTHON_ENTRY_TPL.format(script.script_path.as_posix()))
        for asset in self._assets:
            lines.append("    " + ASSET_ENTRY_TPL.format(asset.path.as_posix()))

        lines.append(MANIFEST_CLOSE_TAG)
        return lines

    def _get_dirs(self) -> List[Path]:
        """
        @return: the destination directories for scripts and assets
        """
        dirs = set()
        for script in self._scripts:
            dirs.update(list(script.script_path.parents)[:-1])
        for asset in self._assets:
            dirs.update(list(asset.path.parents)[:-1])
        return sorted(dirs)
