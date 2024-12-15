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

import io
import unittest
from pathlib import Path
from unittest import mock

from core.asset import SourceAsset, DestinationAsset
from core.script import SourceScript, DestinationScript, TempScript
from core.source_dest import Sources, Destinations
from test.test_helper import file_path_mock

class TestSourcesDests(unittest.TestCase):
    def setUp(self):
        self._sources = Sources(Path("ods_file"), Path("inc"), Path("lib"),
                                Path("src"), [], Path("opt"), Path("asset"), [],
                                Path("test"))
        self._destinations = Destinations(Path("ods_file"), Path("temp"),
                                          Path("dest"), mock.Mock())

    @mock.patch("core.source_dest._get_paths", autospec=True)
    def test_sources_src(self, gp_mock):
        gp_mock.return_value = [Path("a")]

        self.assertEqual([Path("a")], self._sources.get_src_paths())
        self.assertEqual([mock.call(Path('src'), [], '*.py')], gp_mock.mock_calls)

    @mock.patch("core.source_dest._get_paths", autospec=True)
    def test_sources_assets(self, gp_mock):
        gp_mock.return_value = [Path("b")]

        self.assertEqual(
            [SourceAsset(path=Path('b'), assets_dir=Path('asset'))],
            self._sources.get_assets())
        self.assertEqual([mock.call(Path('asset'), [])], gp_mock.mock_calls)

    @mock.patch("core.source_dest._get_paths", autospec=True)
    def test_sources_scripts(self, gp_mock):
        gp_mock.return_value = [Path("oTextRange")]

        self.assertEqual(
            [SourceScript(script_path=Path('oTextRange'), source_dir=Path('src'),
                          export_funcs=True)],
            self._sources.get_src_scripts())
        self.assertEqual([mock.call(Path('src'), [], '*.py')], gp_mock.mock_calls)

    def test_destinations_scripts(self):
        dscript = DestinationScript(Path('dest/scr'), b'abc', Path('dest'),
                                    [], None)
        self.assertEqual([dscript],
                         self._destinations.to_destination_scripts(
                             [TempScript(Path('temp/scr'), b'abc', Path('temp'),
                                         [], None)]))

    def test_destinations_assets(self):
        path = file_path_mock(io.BytesIO(b'asset content'))
        assets_dir: Path = mock.Mock()

        self._destinations.assets_dest_dir.joinpath.return_value = Path(
            "dest_s")

        self.assertEqual(
            [DestinationAsset(path=Path('dest_s'), content=b'asset content')],
            self._destinations.to_destination_assets(
                [SourceAsset(path, assets_dir)]))


if __name__ == '__main__':
    unittest.main()
