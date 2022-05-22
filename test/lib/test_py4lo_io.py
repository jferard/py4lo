#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2022 J. Férard <https://github.com/jferard>
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
from py4lo_io import (create_import_filter_options,
                      create_export_filter_options, Format)

class Py4LOIOTestCase(unittest.TestCase):
    def test_empty_import_options(self):
        self.assertEqual('44,34,76,1,,1033,false,false', create_import_filter_options(language_code="en_US"))

    def test_import_options(self):
        self.assertEqual('44,34,76,1,1/2,1033,false,false',
                         create_import_filter_options(language_code="en_US", format_by_idx={1: Format.TEXT}))

    def test_empty_export_options(self):
        self.assertEqual('44,34,76,1,,1033,false,false,true',
                         create_export_filter_options(language_code="en_US"))


if __name__ == '__main__':
    unittest.main()
