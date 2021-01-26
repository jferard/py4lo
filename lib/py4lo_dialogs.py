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
from threading import Thread

import py4lo_helper


class ProgressExecutorBuilder:
    def __init__(self):
        self._oDialogModel = py4lo_helper.uno_service(
            "com.sun.star.awt.UnoControlDialogModel")
        self._oBar = self._oDialogModel.createInstance(
            "com.sun.star.awt.UnoControlProgressBarModel")
        self._oDialog = py4lo_helper.uno_service(
            "com.sun.star.awt.UnoControlDialog")
        self.title("Please wait...")
        self.dialog_rectangle(150, 150, 150, 30)
        self.bar_dimensions(140, 12)
        self.bar_progress(0, 100)

    def build(self):
        self._oDialogModel.insertByName("bar", self._oBar)
        self._oDialog.setModel(self._oDialogModel)
        self._oDialog.setVisible(True)
        return ProgressExecutor(self._oDialog, self._oBar)

    def title(self, title):
        """
        Set the title
        @param title: the title

        @return: self for fluent style
        """
        self._oDialogModel.Title = title
        return self

    def dialog_rectangle(self, x, y, w, h):
        """
        Set the dialog rectangle

        @param x: abscissa
        @param y: ordinate
        @param w: width
        @param h: height
        @return: self for fluent style
        """
        self._set_rectangle(self._oDialogModel, x, y, w, h)
        return self

    def bar_dimensions(self, w, h):
        """
        Set the dialog rectangle

        @param w: width
        @param h: height
        @return: self for fluent style
        """
        self._set_rectangle(self._oBar, 0, 10, w, h)
        return self

    def bar_progress(self, progress_min, progress_max):
        """
        Set the dialog rectangle

        @param progress_min: min
        @param progress_max: max
        @return: self for fluent style
        """
        self._oBar.ProgressValueMin = progress_min
        self._oBar.ProgressValueMax = progress_max
        return self

    @staticmethod
    def _set_rectangle(o, x, y, w, h):
        o.PositionX = x
        o.PositionY = y
        o.Width = w
        o.Height = h


class ProgressExecutor:
    def __init__(self, oDialog, oBar):
        self._oDialog = oDialog
        self._oBar = oBar

    def execute(self, func):
        """
        Execute the function with a progress bar
        @param func: a function that takes a `ProgressDialog` object.
        """
        def aux():
            func(self._oBar)
            self._oDialog.setVisible(False)
            self._oDialog.dispose()

        t = Thread(target=aux)
        t.start()
