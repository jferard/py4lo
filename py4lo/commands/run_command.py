# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2018 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   Py4LO is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   Py4LO is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""

from commands.update_command import UpdateCommand
from commands.command_executor import CommandExecutor
from tools import open_with_calc
import logging

class RunCommand():
    @staticmethod
    def create(args, tdata):
        update_executor = UpdateCommand.create(args, tdata)
        run_command = RunCommand(tdata["calc_exe"])
        return CommandExecutor(run_command, update_executor)

    def __init__(self, calc_exe):
        self.__calc_exe = calc_exe

    def execute(self, status, dest_name):
        if status == 0:
            print ("All tests ok")
            logging.warning(str(self.__calc_exe)+" "+str(dest_name))
            open_with_calc(dest_name, self.__calc_exe)
        else:
            print ("Error: some tests failed")

    @staticmethod
    def get_help():
        return "Update + open the created file"
