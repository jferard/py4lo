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
import uno

MARGIN = 5
Rectangle = namedtuple('Rectangle', ['x', 'y', 'w', 'h'])
Size = namedtuple('Size', ['w', 'h'])
Progress = namedtuple('Progress', ['min', 'max'])


class ProgressExecutorBuilder:
    def __init__(self):
        self._oDialogModel = py4lo_helper.uno_service(
            "com.sun.star.awt.UnoControlDialogModel")
        self._oBarModel = self._oDialogModel.createInstance(
            "com.sun.star.awt.UnoControlProgressBarModel")
        self._oTextModel = self._oDialogModel.createInstance(
            "com.sun.star.awt.UnoControlFixedTextModel")
        self._oDialog = py4lo_helper.uno_service(
            "com.sun.star.awt.UnoControlDialog")
        self.title("Please wait...")
        self._dialog_rectangle = Rectangle(150, 150, 150, 30)
        self._bar_dimensions = Size(140, 12)
        self._bar_progress = Progress(0, 100)
        self._message = None
        self._autoclose = True

    def build(self):
        self._oDialogModel.insertByName("bar", self._oBarModel)
        self._oDialogModel.insertByName("text", self._oTextModel)
        _set_rectangle(self._oDialogModel, self._dialog_rectangle)
        self._oDialog.setModel(self._oDialogModel)
        x = self._centered(self._dialog_rectangle.w, self._bar_dimensions.w)
        _set_rectangle(self._oBarModel,
                       Rectangle(x, MARGIN, self._bar_dimensions.w,
                                 self._bar_dimensions.h))
        self._oBarModel.ProgressValueMin = self._bar_progress.min
        self._oBarModel.ProgressValueMax = self._bar_progress.max
        self._oBarModel.ProgressValue = self._bar_progress.min
        y = self._bar_dimensions.h + MARGIN * 2
        w = self._dialog_rectangle.w - MARGIN * 2
        h = self._dialog_rectangle.h - self._bar_dimensions.h - MARGIN * 2
        _set_rectangle(self._oTextModel, Rectangle(MARGIN, y, w, h))
        if self._message is not None:
            self._oTextModel.Label = self._message

        return ProgressExecutor(self._oDialog, self._autoclose,
                                self._oDialog.getControl("bar"),
                                self._oDialog.getControl("text"))

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

    def autoclose(self, b):
        """
        Set the autoclose value : if False, don't close
        @param b: the value of autoclose
        @return:
        """
        self._autoclose = b
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
        """
        Set the message
        @param message: the message
        @return: self for fluent style
        """
        self._message = message
        return self


def _set_rectangle(o, rectangle):
    o.PositionX = rectangle.x
    o.PositionY = rectangle.y
    o.Width = rectangle.w
    o.Height = rectangle.h


class ProgressHandler:
    def __init__(self, oBar, oText):
        self._oBar = oBar
        self._oBar.setValue(0)
        self._oText = oText
        self.response = None

    def progress(self, n=1):
        self._oBar.setValue(self._oBar.getValue() + n)

    def message(self, text):
        self._oText.setText(text)


class ProgressExecutor:
    def __init__(self, oDialog, autoclose, oBar, oText):
        self._oDialog = oDialog
        self._autoclose = autoclose
        self._progress_handler = ProgressHandler(oBar, oText)

    def execute(self, func):
        """
        Execute the function with a progress bar
        @param func: a function that takes a `ProgressDialog` object.
        """

        def aux():
            self._oDialog.setVisible(True)
            func(self._progress_handler)
            if self._autoclose:
                self._oDialog.dispose()
            else:
                self._oDialog.execute()  # blocking

        t = Thread(target=aux)
        t.start()

    @property
    def response(self):
        """
        @return: the response
        """
        return self._progress_handler.response


class ConsoleExecutorBuilder:
    def __init__(self):
        self._oDialogModel = py4lo_helper.uno_service(
            "com.sun.star.awt.UnoControlDialogModel")
        self._oDialogModel.Closeable = True
        self._oDialog = py4lo_helper.uno_service(
            "com.sun.star.awt.UnoControlDialog")
        self._oTextModel = self._oDialogModel.createInstance(
            "com.sun.star.awt.UnoControlEditModel")
        self._oTextModel.ReadOnly = True
        self._oTextModel.MultiLine = True
        self._oTextModel.VScroll = True
        self._autoclose = False
        self._console_rectangle = Rectangle(100, 150, 250, 100)
        self.title("Console")

    def title(self, title):
        """
        Set the title
        @param title: the title

        @return: self for fluent style
        """
        self._oDialogModel.Title = title
        return self

    def autoclose(self, b):
        """
        Set the autoclose value : if False, don't close
        @param b: the value of autoclose
        @return:
        """
        self._autoclose = b
        return self

    def console_rectangle(self, x, y, w, h):
        """
        Set the dialog rectangle

        @param x: abscissa
        @param y: ordinate
        @param w: width
        @param h: height
        @return: self for fluent style
        """
        self._console_rectangle = Rectangle(x, y, w, h)
        return self

    def build(self):
        self._oDialog.setModel(self._oDialogModel)
        self._oDialogModel.insertByName("text", self._oTextModel)
        _set_rectangle(self._oDialogModel, self._console_rectangle)
        _set_rectangle(self._oTextModel, Rectangle(MARGIN, MARGIN,
                                                   self._console_rectangle.w - MARGIN * 2,
                                                   self._console_rectangle.h - MARGIN * 2))
        return ConsoleExecutor(self._oDialog, self._autoclose,
                               self._oDialog.getControl("text"))


class ConsoleHandler:
    def __init__(self, oText):
        self._oText = oText
        self.response = None
        self._cur_pos = 0
        self._selection = uno.createUnoStruct('com.sun.star.awt.Selection')

    def message(self, text):
        """
        Add a message to the console
        @param text: the text of the message
        """
        self._selection.Min = self._cur_pos
        self._selection.Max = self._cur_pos
        self._oText.insertText(self._selection, text + "\n")
        self._cur_pos += len(text) + 1


class ConsoleExecutor:
    def __init__(self, oDialog, autoclose, oText):
        self._oDialog = oDialog
        self._autoclose = autoclose
        self._console_handler = ConsoleHandler(oText)

    def execute(self, func):
        """
        Execute the function with a progress bar
        @param func: a function that takes a `ProgressDialog` object.
        """

        def aux():
            self._oDialog.setVisible(True)
            func(self._console_handler)
            if self._autoclose:
                self._oDialog.dispose()
            else:
                self._oDialog.execute()  # blocking

        t = Thread(target=aux)
        t.start()

    @property
    def response(self):
        return self._console_handler.response
