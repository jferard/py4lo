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
import unittest
from logging import Logger
from pathlib import Path
from unittest import mock
from unittest.mock import Mock, MagicMock, patch

from commands.ods_updater import OdsUpdaterHelper
from core.asset import SourceAsset
from core.properties import Sources, Destinations
from core.script import TempScript


class TestOdsUpdaterHelper(unittest.TestCase):
    @patch("py_compile.compile", autospec=True)
    def test(self, compile_mock):
        compile_mock = MagicMock()
        logger: Logger = Mock()
        sources: Sources = Mock(src_ignore=["*.pyc"])
        destinations: Destinations = Mock()
        destinations.temp_dir.joinpath.side_effect = [Path("temp.py")]
        td: Path = Mock()
        sp: Path = Mock()
        bound = MagicMock()

        sp.relative_to.return_value = Path("reldest")
        sp.open.side_effect = bound
        td.joinpath.side_effect = [sp]
        destinations.temp_dir = td
        bound.__enter__.return_value = ["print('hello')"]
        path: Path = Mock()

        path.is_file.return_value = True
        path.open.side_effect = [bound]
        sources.src_dir.rglob.side_effect = [[path], []]
        asset: SourceAsset = Mock()
        asset.to_dest.return_value = 0
        sources.get_assets.return_value = [asset]

        h = OdsUpdaterHelper(logger, sources, destinations, "3.1")
        actual = h.get_temp_scripts()
        self.assertEqual([TempScript(sp, b"# parsed by py4lo (https://github.com/jferard/py4lo)\nprint('hello')",
                                     destinations.temp_dir, [], None)],
                         actual)

        self.assertEqual([0], h.get_assets())
