#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2025 J. Férard <https://github.com/jferard>
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
from unittest import mock
from zipfile import ZipFile

from callbacks import AddAssets
from core.asset import DestinationAsset


class AddAssetsTest(unittest.TestCase):
    def test_empty(self):
        zout: ZipFile = mock.Mock()
        cb = AddAssets([])
        cb.call(zout)

        self.assertEqual([], zout.mock_calls)

    def test_empty2(self):
        zout: ZipFile = mock.Mock()
        asset: DestinationAsset = mock.Mock(path=Path("asset"), content=bytes())

        cb = AddAssets([asset])
        cb.call(zout)

        self.assertEqual([mock.call.writestr('asset', b'')],
                         zout.mock_calls)
