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

# see http://stackoverflow.com/questions/61151/where-do-the-python-unit-tests-go
import sys
from pathlib import Path

def any_object():
    class AnyObject:
        def __eq__(self, other):
            return True
    return AnyObject()


# append module root directory to sys.path
test_dir = Path(__file__).parent
root_dir = test_dir.parent
py4lo_dir = root_dir.joinpath("py4lo")
lib_dir = root_dir.joinpath("lib")
inc_dir = root_dir.joinpath("inc")

for p in map(str, [py4lo_dir, lib_dir, inc_dir]):
    if p not in sys.path:
        sys.path.insert(0, p)