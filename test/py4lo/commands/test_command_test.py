# -*- coding: utf-8 -*-
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

import subprocess   # nosec: B404
import unittest
from pathlib import Path
from unittest import mock

from commands.test_command import TestCommand


class TestCommandTest(unittest.TestCase):
    @mock.patch('subprocess.run', spec=subprocess.run)
    def test(self, subprocess_run_mock):
        completed_process1 = mock.MagicMock(returncode=0)
        completed_process1.stdout.decode.return_value = "ok"
        completed_process2 = mock.MagicMock(returncode=1)
        completed_process2.stdout.decode.return_value = "not ok"
        completed_process2.stderr.decode.return_value = "err"
        completed_process3 = mock.MagicMock(returncode=0)
        completed_process3.stdout.decode.return_value = "s ok"
        subprocess_run_mock.side_effect = [completed_process1,
                                           completed_process2,
                                           completed_process3]
        logger = mock.MagicMock()
        sources = mock.MagicMock()
        sources.test_dir.rglob.side_effect = [[Path("/test_dir/c_test.py"),
                                               Path("/test_dir/b/d_test.py")]]
        sources.src_dir.__truediv__.side_effect = [[Path("/src_dir/main.py")]]
        sources.src_dir.rglob.side_effect = [[Path("/src_dir/src_a.py")]]
        tc = TestCommand(logger, "test_py_exe", sources)
        status = tc.execute()

        self.assertEqual([
            mock.call.info(
                'execute doctests: %s',
                '"test_py_exe" -m doctest /src_dir/src_a.py'),
            mock.call.debug('PYTHONPATH = %s', mock.ANY),
            mock.call.info('output: ok'),
            mock.call.info(
                'execute unittests: "test_py_exe" /test_dir/c_test.py'),
            mock.call.info('output: not ok'),
            mock.call.error('error: err'),
            mock.call.info(
                'execute unittests: "test_py_exe" /test_dir/b/d_test.py'),
            mock.call.info('output: s ok'),
        ], logger.mock_calls)
        self.assertEqual([
            mock.call(["test_py_exe", "-m", "doctest", "/src_dir/src_a.py"],
                      env=unittest.mock.ANY, stderr=-1, stdout=-1),
            mock.call(["test_py_exe", "/test_dir/c_test.py"],
                      env=mock.ANY, stderr=-1, stdout=-1),
            mock.call(["test_py_exe", "/test_dir/b/d_test.py"],
                      env=mock.ANY, stderr=-1, stdout=-1),
        ], subprocess_run_mock.mock_calls)
        self.assertEqual((1,), status)


if __name__ == '__main__':
    unittest.main()
