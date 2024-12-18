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
import logging
import unittest
from unittest import mock
from zipfile import ZipFile
from pathlib import Path

from core.properties import PropertiesProviderFactory, PropertiesProvider
from core.source_dest import _get_paths


class TestProperties(unittest.TestCase):
    @mock.patch("core.properties.load_toml", autospec=True)
    def test_provider_factory(self, toml_mock):
        toml_mock.return_value = {
            "a": 1, "log_level": -1,
            "src": {"source_ods_file": "s.ods",
                    "inc_dir": "inc", "lib_dir": "lib",
                    "src_dir": "src", "src_ignore": "",
                    "opt_dir": "opt", "assets_dir": "assets",
                    "assets_ignore": "", "test_dir": "test"},
            "dest": {"suffix": "ok", "temp_dir": "temp",
                     "dest_dir": "dest", "assets_dest_dir": "Assets"}
        }
        p = PropertiesProviderFactory().create()

        self.assertEqual({"a", "log_level", "src", "dest"}, p.keys())
        self.assertEqual(1, p.get("a"))

    def test_provider(self):
        logger = logging.getLogger()
        sources = mock.Mock(get_assets_paths=lambda: Path("assets"),
                            get_src_paths=lambda: Path("src"))
        destinations = mock.Mock()
        provider = PropertiesProvider(logger, Path("base"), sources,
                                      destinations, {'a': 1})
        self.assertEqual(logger, provider.get_logger())
        self.assertEqual(1, provider.get('a'))
        self.assertEqual(None, provider.get('b'))
        self.assertEqual(10, provider.get('b', 10))
        self.assertEqual(Path("base"), provider.get_base_path())
        self.assertEqual(sources, provider.get_sources())
        self.assertEqual(Path("assets"), provider.get_assets_paths())
        self.assertEqual(Path("src"), provider.get_src_paths())
        self.assertEqual(destinations, provider.get_destinations())
        self.assertEqual(None, provider.get_readme_callback())

    def test_readme(self):
        logger = logging.getLogger()
        sources = mock.Mock(get_assets_paths=lambda: Path("assets"),
                            get_src_paths=lambda: Path("src"))
        destinations = mock.Mock()
        self.maxDiff = None
        base_path = Path(__file__).parent.parent.parent.parent
        provider = PropertiesProvider(logger, base_path, sources,
                                      destinations, {'add_readme': True})
        f = io.BytesIO()
        with ZipFile(f, "w") as zout:
            provider.get_readme_callback().call(zout)
        with ZipFile(f) as zin:
            self.assertEqual(['Basic/script-lc.xml',
                              'Basic/Standard/script-lb.xml',
                              'Basic/Standard/py4lo.xml'], zin.namelist())

    def test_get_paths(self):
        source_dir: Path = mock.Mock()
        path_a: Path = mock.Mock()
        path_b: Path = mock.Mock()

        source_dir.rglob.side_effect = [[path_a, path_b, Path("./oTextRange")],
                                        [Path("oTextRange")]]
        path_a.is_file.return_value = True
        path_b.is_file.return_value = False

        self.assertEqual({path_a},
                         _get_paths(source_dir, ["oTextRange"], "*"))
        self.assertEqual([mock.call.rglob('*'), mock.call.rglob('oTextRange')],
                         source_dir.mock_calls)


if __name__ == '__main__':
    unittest.main()
