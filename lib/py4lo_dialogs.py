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
from enum import Enum
from threading import Thread
from typing import Any, Callable, Optional, List, Union, NamedTuple

from collections import namedtuple


from py4lo_commons import StrPath
from py4lo_helper import (uno_service_ctxt, provider, uno_service)
from py4lo_typing import UnoObject, UnoControlModel, UnoControl

try:
    import uno

    class MessageBoxType:
        from com.sun.star.awt.MessageBoxType import (ERRORBOX, MESSAGEBOX)


    class MessageBoxButtons:
        from com.sun.star.awt.MessageBoxButtons import (BUTTONS_OK, )


    class FontWeight:
        from com.sun.star.awt.FontWeight import (BOLD, )


    class ExecutableDialogResults:
        from com.sun.star.ui.dialogs.ExecutableDialogResults import (OK, CANCEL)
except ModuleNotFoundError:
    class MessageBoxType:
        MESSAGEBOX = None


    class MessageBoxButtons:
        BUTTONS_OK = None


class ControlModel(str, Enum):
    AnimatedImages = "AnimatedImagesControlModel"
    Grid = "UnoControlGridModel"
    TabPageContainer = "UnoControlTabPageContainerModel"
    Tree = "TreeControlModel"
    Button = "UnoControlButtonModel"
    CheckBox = "UnoControlCheckBoxModel"
    ComboBox = "UnoControlComboBoxModel"
    Container = "UnoControlContainerModel"
    CurrencyField = "UnoControlCurrencyFieldModel"
    DateField = "UnoControlDateFieldModel"
    Dialog = "UnoControlDialogModel"
    Edit = "UnoControlEditModel"
    FileControl = "UnoControlFileControlModel"
    FixedHyperlink = "UnoControlFixedHyperlinkModel"
    FixedLine = "UnoControlFixedLineModel"
    FixedText = "UnoControlFixedTextModel"
    FormattedField = "UnoControlFormattedFieldModel"
    GroupBox = "UnoControlGroupBoxModel"
    ImageControl = "UnoControlImageControlModel"
    ListBox = "UnoControlListBoxModel"
    NumericField = "UnoControlNumericFieldModel"
    PatternField = "UnoControlPatternFieldModel"
    ProgressBar = "UnoControlProgressBarModel"
    RadioButton = "UnoControlRadioButtonModel"
    Roadmap = "UnoControlRoadmapModel"
    ScrollBar = "UnoControlScrollBarModel"
    SpinButton = "UnoControlSpinButtonModel"
    TimeField = "UnoControlTimeFieldModel"
    ColumnDescriptor = "ColumnDescriptorControlModel"


class Control(str, Enum):
    AnimatedImages = "AnimatedImagesControl"
    Grid = "UnoControlGrid"
    TabPageContainer = "UnoControlTabPageContainer"
    Tree = "TreeControl"
    Button = "UnoControlButton"
    CheckBox = "UnoControlCheckBox"
    ComboBox = "UnoControlComboBox"
    Container = "UnoControlContainer"
    CurrencyField = "UnoControlCurrencyField"
    DateField = "UnoControlDateField"
    Dialog = "UnoControlDialog"
    Edit = "UnoControlEdit"
    FileControl = "UnoControlFileControl"
    FixedHyperlink = "UnoControlFixedHyperlink"
    FixedLine = "UnoControlFixedLine"
    FixedText = "UnoControlFixedText"
    FormattedField = "UnoControlFormattedField"
    GroupBox = "UnoControlGroupBox"
    ImageControl = "UnoControlImageControl"
    ListBox = "UnoControlListBox"
    NumericField = "UnoControlNumericField"
    PatternField = "UnoControlPatternField"
    ProgressBar = "UnoControlProgressBar"
    RadioButton = "UnoControlRadioButton"
    Roadmap = "UnoControlRoadmap"
    ScrollBar = "UnoControlScrollBar"
    SpinButton = "UnoControlSpinButton"
    TimeField = "UnoControlTimeField"
    ColumnDescriptor = "ColumnDescriptorControl"


###
# Common functions
###

def place_widget(
        oWidgetModel: UnoControlModel, x: int, y: int, width: int, height: int):
    """Place a widget on the main model"""
    oWidgetModel.PositionX = x
    oWidgetModel.PositionY = y
    oWidgetModel.Width = width
    oWidgetModel.Height = height


def message_box(msg_text: str, msg_title: str,
                msg_type=MessageBoxType.MESSAGEBOX,
                msg_buttons=MessageBoxButtons.BUTTONS_OK, parent_win=None):
    """Create a message box"""
    # from https://forum.openoffice.org/fr/forum/viewtopic.php?f=15&t=47603#
    # (thanks Bernard !)
    toolkit = uno_service_ctxt("com.sun.star.awt.Toolkit")
    if parent_win is None:
        parent_win = provider.parent_win
    mb = toolkit.createMessageBox(parent_win, msg_type, msg_buttons, msg_title,
                                  msg_text)
    return mb.execute()


FileFilter = NamedTuple("FileFilter", [("title", str), ("filter", str)])


def file_dialog(title: str, filters: Optional[List[FileFilter]] = None,
                display_dir: StrPath = "",
                single: bool = True) -> Union[Optional[str], List[str]]:
    """
    Open a file dialog
    @return: if single, url or None, else a list of urls
    """
    oFilePicker = uno_service("com.sun.star.ui.dialogs.FilePicker")
    if filters is not None:
        for flt in filters:
            oFilePicker.appendFilter(flt.title, flt.filter)

    oFilePicker.Title = title
    oFilePicker.DisplayDirectory = str(display_dir)
    if single:
        oFilePicker.MultiSelectionMode = False
        if oFilePicker.execute() == ExecutableDialogResults.OK:
            urls = oFilePicker.getSelectedFiles()
            return urls[0] if urls else None
        else:
            return None
    else:
        oFilePicker.MultiSelectionMode = True
        if oFilePicker.execute() == ExecutableDialogResults.OK:
            return oFilePicker.getSelectedFiles()
        else:
            return []


def folder_dialog(title: str,
                  display_dir: StrPath = "") -> Optional[str]:
    """
    Open a file dialog
    @return: url or None
    """
    oFolder = uno_service("com.sun.star.ui.dialogs.FolderPicker")
    oFolder.Title = title
    oFolder.DisplayDirectory = str(display_dir)
    if oFolder.execute() == ExecutableDialogResults.OK:
        return oFolder.Directory
    else:
        return None


MARGIN = 5
Rectangle = namedtuple('Rectangle', ['x', 'y', 'w', 'h'])
Size = namedtuple('Size', ['w', 'h'])
Progress = namedtuple('Progress', ['min', 'max'])


class ProgressExecutorBuilder:
    def __init__(self):
        self._oDialogModel = uno_service(ControlModel.Dialog)
        self._oBarModel = self._oDialogModel.createInstance(
            ControlModel.ProgressBar)
        self._oTextModel = self._oDialogModel.createInstance(
            ControlModel.FixedText)
        self._oDialog = uno_service(Control.Dialog)
        self.title("Please wait...")
        self._dialog_rectangle = Rectangle(150, 150, 150, 30)
        self._bar_dimensions = Size(140, 12)
        self._bar_progress = Progress(0, 100)
        self._message = None
        self._autoclose = True

    def build(self) -> "ProgressExecutor":
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

    def _centered(self, outer_w: int, inner_w: int) -> int:
        if outer_w <= inner_w:
            return 0
        else:
            return (outer_w - inner_w) // 2

    def title(self, title: str) -> "ProgressExecutorBuilder":
        """
        Set the title
        @param title: the title

        @return: self for fluent style
        """
        self._oDialogModel.Title = title
        return self

    def autoclose(self, b: bool) -> "ProgressExecutorBuilder":
        """
        Set the autoclose value : if False, don't close
        @param b: the value of autoclose
        @return:
        """
        self._autoclose = b
        return self

    def dialog_rectangle(self, x: int, y: int, w: int,
                         h: int) -> "ProgressExecutorBuilder":
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

    def bar_progress(self, progress_min: int,
                     progress_max: int) -> "ProgressExecutorBuilder":
        """
        Set the dialog rectangle

        @param progress_min: min
        @param progress_max: max
        @return: self for fluent style
        """
        self._bar_progress = Progress(progress_min, progress_max)
        return self

    def message(self, message: str) -> "ProgressExecutorBuilder":
        """
        Set the message
        @param message: the message
        @return: self for fluent style
        """
        self._message = message
        return self


def _set_rectangle(o: Any, rectangle: Rectangle):
    o.PositionX = rectangle.x
    o.PositionY = rectangle.y
    o.Width = rectangle.w
    o.Height = rectangle.h


class ProgressHandler:
    def __init__(self, oBar: UnoControlModel, oText: UnoControlModel):
        self._oBar = oBar
        self._oBar.setValue(0)
        self._oText = oText
        self.response = None

    def progress(self, n: int = 1):
        self._oBar.setValue(self._oBar.getValue() + n)

    def message(self, text: str):
        self._oText.setText(text)


class ProgressExecutor:
    def __init__(self, oDialog: UnoControl, autoclose: bool,
                 oBar: UnoControlModel, oText: UnoControlModel):
        self._oDialog = oDialog
        self._autoclose = autoclose
        self._progress_handler = ProgressHandler(oBar, oText)

    def execute(self, func: Callable[[ProgressHandler], None]):
        """
        Execute the function with a progress bar
        @param func: a function that takes a `ProgressDialog` object.
        """

        def aux():
            # TODO:     oDialogControl.setVisible(True)
            #     toolkit = uno_service("com.sun.star.awt.Toolkit")
            #     oDialogControl.createPeer(toolkit, None)
            self._oDialog.setVisible(True)
            func(self._progress_handler)
            if self._autoclose:
                self._oDialog.dispose()
            else:
                self._oDialog.execute()  # blocking

        t = Thread(target=aux)
        t.start()

    @property
    def response(self) -> Any:
        """
        @return: the response
        """
        return self._progress_handler.response


class ConsoleExecutorBuilder:
    def __init__(self):
        self._oDialogModel = uno_service(ControlModel.Dialog)
        self._oDialogModel.Closeable = True
        self._oDialog = uno_service(Control.Dialog)
        self._oTextModel = self._oDialogModel.createInstance(ControlModel.Edit)
        self._oTextModel.ReadOnly = True
        self._oTextModel.MultiLine = True
        self._oTextModel.VScroll = True
        self._autoclose = False
        self._console_rectangle = Rectangle(100, 150, 250, 100)
        self.title("Console")

    def title(self, title: str) -> "ConsoleExecutorBuilder":
        """
        Set the title
        @param title: the title

        @return: self for fluent style
        """
        self._oDialogModel.Title = title
        return self

    def autoclose(self, b: bool) -> "ConsoleExecutorBuilder":
        """
        Set the autoclose value : if False, don't close
        @param b: the value of autoclose
        @return:
        """
        self._autoclose = b
        return self

    def console_rectangle(self, x: int, y: int, w: int,
                          h: int) -> "ConsoleExecutorBuilder":
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

    def build(self) -> "ConsoleExecutor":
        self._oDialog.setModel(self._oDialogModel)
        self._oDialogModel.insertByName("text", self._oTextModel)
        _set_rectangle(self._oDialogModel, self._console_rectangle)
        _set_rectangle(self._oTextModel, Rectangle(MARGIN, MARGIN,
                                                   self._console_rectangle.w - MARGIN * 2,
                                                   self._console_rectangle.h - MARGIN * 2))
        return ConsoleExecutor(self._oDialog, self._autoclose,
                               self._oDialog.getControl("text"))


class ConsoleHandler:
    def __init__(self, oText: UnoControlModel):
        self._oText = oText
        self.response = None
        self._cur_pos = 0
        self._selection = uno.createUnoStruct('com.sun.star.awt.Selection')

    def message(self, text: str):
        """
        Add a message to the console
        @param text: the text of the message
        """
        self._selection.Min = self._cur_pos
        self._selection.Max = self._cur_pos
        self._oText.insertText(self._selection, text + "\n")
        self._cur_pos += len(text) + 1


class ConsoleExecutor:
    def __init__(self, oDialog: UnoControl, autoclose: bool,
                 oText: UnoControlModel):
        self._oDialog = oDialog
        self._autoclose = autoclose
        self._console_handler = ConsoleHandler(oText)

    def execute(self, func):
        """
        Execute the function with a progress bar
        @param func: a function that takes a `ProgressDialog` object.
        """

        def aux():
            # TODO:     oDialogControl.setVisible(True)
            #     toolkit = uno_service("com.sun.star.awt.Toolkit")
            #     oDialogControl.createPeer(toolkit, None)
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
