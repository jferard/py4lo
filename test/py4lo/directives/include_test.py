#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
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
from unittest import mock

from directives import Include

from test.test_helper import file_path_mock, verify_open_path


class TestInclude(unittest.TestCase):
    def test_without_strip(self):
        proc = mock.Mock()
        path = mock.Mock()
        inc_path = file_path_mock(io.StringIO("some line"))

        path.joinpath.return_value = inc_path
        inc_path.__str__.return_value = "[to inc]"

        d = Include(path)
        self.assertEqual(["include"], d.sig_elements())
        self.assertEqual(True, d.execute(None, proc, ["a.py"]))
        self.assertEqual(
            [
                mock.call.append(
                    "# begin py4lo include: [to inc]\nsome line\n# end py4lo include\n"
                )
            ],
            proc.mock_calls,
        )

        verify_open_path(self, inc_path, "r", encoding="utf-8")

    def test_with_strip(self):
        proc = mock.Mock()
        path = mock.Mock()
        inc_path = file_path_mock(
            io.StringIO(
                "\"\"\"docstring\"\"\"\n'''\nother docstring'''\n  #comment\nsome line"  # noqa: E501
            )
        )

        path.joinpath.return_value = inc_path
        inc_path.__str__.return_value = "[to inc]"

        d = Include(path)
        self.assertEqual(["include"], d.sig_elements())
        self.assertEqual(True, d.execute(None, proc, ["a.py", "True"]))
        self.assertEqual(
            [
                mock.call.append(
                    "# begin py4lo include: [to inc]\nsome line\n# end py4lo include\n"
                )
            ],
            proc.mock_calls,
        )

        verify_open_path(self, inc_path, "r", encoding="utf-8")


if __name__ == "__main__":
    unittest.main()
