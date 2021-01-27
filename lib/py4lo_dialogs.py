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
from collections import namedtuple

MARGIN = 5
Rectangle = namedtuple('Rectangle', ['x', 'y', 'w', 'h'])
Size = namedtuple('Size', ['w', 'h'])
Progress = namedtuple('Progress', ['min', 'max'])


class ProgressExecutorBuilder:
    def __init__(self):
        self._oDialogModel = py4lo_helper.uno_service(
            "com.sun.star.awt.UnoControlDialogModel")
        self._oBar = self._oDialogModel.createInstance(
            "com.sun.star.awt.UnoControlProgressBarModel")
        self._oText = self._oDialogModel.createInstance(
            "com.sun.star.awt.UnoControlFixedTextModel")
        self._oDialog = py4lo_helper.uno_service(
            "com.sun.star.awt.UnoControlDialog")
        self.title("Please wait...")
        self._dialog_rectangle = Rectangle(150, 150, 150, 30)
        self._bar_dimensions = Size(140, 12)
        self._bar_progress = Progress(0, 100)
        self._message = None

    def build(self):
        self._oDialogModel.insertByName("bar", self._oBar)
        self._oDialogModel.insertByName("text", self._oText)
        self._set_rectangle(self._oDialogModel, self._dialog_rectangle)
        self._oDialog.setModel(self._oDialogModel)
        x = self._centered(self._dialog_rectangle.w, self._bar_dimensions.w)
        self._set_rectangle(self._oBar,
                            Rectangle(x, MARGIN, self._bar_dimensions.w,
                                      self._bar_dimensions.h))
        self._oBar.ProgressValueMin = self._bar_progress.min
        self._oBar.ProgressValueMax = self._bar_progress.max
        self._oBar.ProgressValue = self._bar_progress.min
        y = self._bar_dimensions.h + MARGIN * 2
        w = self._dialog_rectangle.w - MARGIN * 2
        h = self._dialog_rectangle.h - self._bar_dimensions.h - MARGIN * 2
        self._set_rectangle(self._oText, Rectangle(MARGIN, y, w, h))
        if self._message is not None:
            self._oText.Label = self._message

        return ProgressExecutor(self._oDialog, self._oBar, self._oText)

    def _centered(self, outer_w, inner_w):
        if outer_w <= inner_w:
            return 0
        else:
            return (outer_w - inner_w) // 2

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
        self._dialog_rectangle = Rectangle(x, y, w, h)
        return self

    def bar_dimensions(self, w, h):
        """
        Set the dialog rectangle

        @param w: width
        @param h: height
        @return: self for fluent style
        """
        self._bar_dimensions = Size(w, h)
        return self

    def bar_progress(self, progress_min, progress_max):
        """
        Set the dialog rectangle

        @param progress_min: min
        @param progress_max: max
        @return: self for fluent style
        """
        self._bar_progress = Progress(progress_min, progress_max)
        return self

    def message(self, message):
        self._message = message

    @staticmethod
    def _set_rectangle(o, rectangle):
        o.PositionX = rectangle.x
        o.PositionY = rectangle.y
        o.Width = rectangle.w
        o.Height = rectangle.h


class ProgressHandler:
    def __init__(self, oBar, oText):
        self._oBar = oBar
        self._oBar.ProgressValue = 0
        self._oText = oText
        self.response = None

    def progress(self, n=1):
        self._oBar.ProgressValue += n

    def message(self, text):
        self._oText.Label = text


class ProgressExecutor:
    def __init__(self, oDialog, oBar, oText):
        self._oDialog = oDialog
        self._progress_handler = ProgressHandler(oBar, oText)

    def execute(self, func):
        """
        Execute the function with a progress bar
        @param func: a function that takes a `ProgressDialog` object.
        """

        def aux():
            self._oDialog.setVisible(True)
            func(self._progress_handler)
            self._oDialog.setVisible(False)
            self._oDialog.dispose()

        t = Thread(target=aux)
        t.start()

    @property
    def response(self):
        return self._progress_handler.response
