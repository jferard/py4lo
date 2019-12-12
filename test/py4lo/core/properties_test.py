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
import io
import unittest
from unittest.mock import patch, call, Mock

import env
from core.properties import *
from core.properties import _get_paths
from env import file_path_mock


class TestProperties(unittest.TestCase):
    def setUp(self):
        self._sources = Sources(Path("ods_file"), Path("inc"), Path("lib"),
                                Path("src"), [], Path("opt"), Path("asset"), [],
                                Path("test"))
        self._destinations = Destinations(Path("ods_file"), Path("temp"),
                                          Path("dest"), Mock())

    @patch("core.properties.load_toml", autospec=True)
    def test_provider(self, toml_mock):
        toml_mock.return_value = {"a": 1, "log_level": -1,
                                  "src": {"source_ods_file": "s.ods",
                                          "inc_dir": "inc", "lib_dir": "lib",
                                          "src_dir": "src", "src_ignore": "",
                                          "opt_dir": "opt",
                                          "assets_dir": "assets",
                                          "assets_ignore": "",
                                          "test_dir": "test"},
                                  "dest": {"suffix": "ok", "temp_dir": "temp",
                                           "dest_dir": "dest",
                                           "assets_dest_dir": "Assets"}}
        p = PropertiesProviderFactory().create()

        self.assertEqual({"a", "log_level", "src", "dest"}, p.keys())
        self.assertEqual(1, p.get("a"))

    def test_get_paths(self):
        source_dir: Path = Mock()
        path_a: Path = Mock()
        path_b: Path = Mock()

        source_dir.rglob.side_effect = [[path_a, path_b, Path("./c")],
                                        [Path("c")]]
        path_a.is_file.return_value = True
        path_b.is_file.return_value = False

        self.assertEqual(set([path_a]),
                         _get_paths(source_dir, ["c"], "*"))
        self.assertEqual([call.rglob('*'), call.rglob('c')],
                         source_dir.mock_calls)

    @patch("core.properties._get_paths", autospec=True)
    def test_sources_src(self, gp_mock):
        gp_mock.return_value = [Path("a")]

        self.assertEqual([Path("a")], self._sources.get_src_paths())
        self.assertEqual([call(Path('src'), [], '*.py')], gp_mock.mock_calls)

    @patch("core.properties._get_paths", autospec=True)
    def test_sources_assets(self, gp_mock):
        gp_mock.return_value = [Path("b")]

        self.assertEqual(
            [SourceAsset(path=Path('b'), assets_dir=Path('asset'))],
            self._sources.get_assets())
        self.assertEqual([call(Path('asset'), [])], gp_mock.mock_calls)

    @patch("core.properties._get_paths", autospec=True)
    def test_sources_scripts(self, gp_mock):
        gp_mock.return_value = [Path("c")]

        self.assertEqual(
            [SourceScript(script_path=Path('c'), source_dir=Path('src'))],
            self._sources.get_src_scripts())
        self.assertEqual([call(Path('src'), [], '*.py')], gp_mock.mock_calls)

    def test_destinations_scripts(self):
        dscript = DestinationScript(Path('dest/scr'), b'abc', Path('dest'),
                                    [], None)
        self.assertEqual([dscript],
                         self._destinations.to_destination_scripts(
                             [TempScript(Path('temp/scr'), b'abc', Path('temp'),
                                         [], None)]))

    def test_destinations_assets(self):
        path = file_path_mock(io.BytesIO(b'asset content'))
        assets_dir: Path = Mock()

        self._destinations.assets_dest_dir.joinpath.return_value = Path("dest_s")

        self.assertEqual([DestinationAsset(path=Path('dest_s'), content=b'asset content')],
                         self._destinations.to_destination_assets(
                             [SourceAsset(path, assets_dir)]))


if __name__ == '__main__':
    unittest.main()
