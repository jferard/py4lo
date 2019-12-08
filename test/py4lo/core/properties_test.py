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
from pathlib import Path
from unittest.mock import patch, call

from core.properties import PropertiesProviderFactory


class TestProperties(unittest.TestCase):
    @patch("core.properties.load_toml", autospec=True)
    def test_provider(self, toml_mock):
        toml_mock.return_value = {"a": 1, "log_level": -1,
                                  "src": {"source_ods_file": "s.ods",
                                          "inc_dir": "inc", "lib_dir": "lib",
                                          "src_dir": "src", "src_ignore": "",
                                          "assets_dir": "assets",
                                          "assets_ignore": "",
                                          "test_dir": "test"},
                                  "dest": {"suffix": "ok", "temp_dir": "temp",
                                           "dest_dir": "dest",
                                           "assets_dest_dir": "Assets"}}
        p = PropertiesProviderFactory().create()

        self.assertEqual({"a", "log_level", "src", "dest"}, p.keys())
        self.assertEqual(1, p.get("a"))


if __name__ == '__main__':
    unittest.main()
