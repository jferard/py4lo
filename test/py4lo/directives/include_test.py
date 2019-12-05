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
from unittest.mock import Mock, call, MagicMock

from directives import Include


class TestInclude(unittest.TestCase):
    def test(self):
        proc = Mock()
        path = Mock()
        inc_path = MagicMock()
        bound = MagicMock()
        inc_path.__str__.return_value = "[to inc]"
        path.joinpath.return_value = inc_path
        inc_path.open.return_value = bound
        bound.__enter__.return_value = ["some line"]

        d = Include(path)
        self.assertEqual(["include"], d.sig_elements())
        self.assertEqual(True, d.execute(proc, ["a.py"]))
        self.assertEqual([call.append(
            '# begin py4lo include: [to inc]\nsome line\n# end py4lo include\n')],
            proc.mock_calls)

        self.assertEqual(
            [call.__getattr__("__str__"), call.open('r', encoding="utf-8"),
             call.open().__enter__(),
             call.open().__exit__(None, None, None)], inc_path.mock_calls)
        self.assertEqual([call.__enter__(), call.__exit__(None, None, None)],
                         bound.mock_calls)


if __name__ == '__main__':
    unittest.main()
