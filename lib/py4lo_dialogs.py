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

from py4lo_helper import (uno_service_ctxt, provider, uno_service)
from py4lo_typing import UnoControlModel, UnoControl, StrPath

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
    AnimatedImages = "com.sun.star.awt.AnimatedImagesControlModel"
    Grid = "com.sun.star.awt.grid.UnoControlGridModel"
    TabPageContainer = "com.sun.star.awt.tab.UnoControlTabPageContainerModel"
    Tree = "com.sun.star.awt.tree.TreeControlModel"
    Button = "com.sun.star.awt.UnoControlButtonModel"
    CheckBox = "com.sun.star.awt.UnoControlCheckBoxModel"
    ComboBox = "com.sun.star.awt.UnoControlComboBoxModel"
    Container = "com.sun.star.awt.UnoControlContainerModel"
    CurrencyField = "com.sun.star.awt.UnoControlCurrencyFieldModel"
    DateField = "com.sun.star.awt.UnoControlDateFieldModel"
    Dialog = "com.sun.star.awt.UnoControlDialogModel"
    Edit = "com.sun.star.awt.UnoControlEditModel"
    FileControl = "com.sun.star.awt.UnoControlFileControlModel"
    FixedHyperlink = "com.sun.star.awt.UnoControlFixedHyperlinkModel"
    FixedLine = "com.sun.star.awt.UnoControlFixedLineModel"
    FixedText = "com.sun.star.awt.UnoControlFixedTextModel"
    FormattedField = "com.sun.star.awt.UnoControlFormattedFieldModel"
    GroupBox = "com.sun.star.awt.UnoControlGroupBoxModel"
    ImageControl = "com.sun.star.awt.UnoControlImageControlModel"
    ListBox = "com.sun.star.awt.UnoControlListBoxModel"
    NumericField = "com.sun.star.awt.UnoControlNumericFieldModel"
    PatternField = "com.sun.star.awt.UnoControlPatternFieldModel"
    ProgressBar = "com.sun.star.awt.UnoControlProgressBarModel"
    RadioButton = "com.sun.star.awt.UnoControlRadioButtonModel"
    Roadmap = "com.sun.star.awt.com.sun.star.awt.UnoControlRoadmapModel"
    ScrollBar = "com.sun.star.awt.UnoControlScrollBarModel"
    SpinButton = "com.sun.star.awt.UnoControlSpinButtonModel"
    TimeField = "com.sun.star.awt.UnoControlTimeFieldModel"
    ColumnDescriptor = "com.sun.star.sdb.ColumnDescriptorControlModel"


class Control(str, Enum):
    AnimatedImages = "com.sun.star.awt.AnimatedImagesControl"
    Grid = "com.sun.star.awt.grid.UnoControlGrid"
    TabPageContainer = "com.sun.star.awt.tab.UnoControlTabPageContainer"
    Tree = "com.sun.star.awt.tree.TreeControl"
    Button = "com.sun.star.awt.UnoControlButton"
    CheckBox = "com.sun.star.awt.UnoControlCheckBox"
    ComboBox = "com.sun.star.awt.UnoControlComboBox"
    Container = "com.sun.star.awt.UnoControlContainer"
    CurrencyField = "com.sun.star.awt.UnoControlCurrencyField"
    DateField = "com.sun.star.awt.UnoControlDateField"
    Dialog = "com.sun.star.awt.UnoControlDialog"
    Edit = "com.sun.star.awt.UnoControlEdit"
    FileControl = "com.sun.star.awt.UnoControlFileControl"
    FixedHyperlink = "com.sun.star.awt.UnoControlFixedHyperlink"
    FixedLine = "com.sun.star.awt.UnoControlFixedLine"
    FixedText = "com.sun.star.awt.UnoControlFixedText"
    FormattedField = "com.sun.star.awt.UnoControlFormattedField"
    GroupBox = "com.sun.star.awt.UnoControlGroupBox"
    ImageControl = "com.sun.star.awt.UnoControlImageControl"
    ListBox = "com.sun.star.awt.UnoControlListBox"
    NumericField = "com.sun.star.awt.UnoControlNumericField"
    PatternField = "com.sun.star.awt.UnoControlPatternField"
    ProgressBar = "com.sun.star.awt.UnoControlProgressBar"
    RadioButton = "com.sun.star.awt.UnoControlRadioButton"
    Roadmap = "com.sun.star.awt.UnoControlRoadmap"
    ScrollBar = "com.sun.star.awt.UnoControlScrollBar"
    SpinButton = "com.sun.star.awt.UnoControlSpinButton"
    TimeField = "com.sun.star.awt.UnoControlTimeField"
    ColumnDescriptor = "com.sun.star.sdb.ColumnDescriptorControl"


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


Size = namedtuple("Size", ["width", "height"])


def get_text_size(oDialogModel: UnoControlModel, text: str) -> Size:
    """
    @param oDialogModel: the model
    @param text: the text
    @return: the text size
    """
    oTextControl = uno_service(Control.FixedText)
    oTextModel = oDialogModel.createInstance(ControlModel.FixedText)
    oTextModel.Label = text
    oTextControl.Model = oTextModel
    min_size = oTextControl.MinimumSize
    # Why 0.5 ? I don't know
    return Size(min_size.Width * 0.5, min_size.Height * 0.5)


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
            urls = oFilePicker.SelectedFiles
            return urls[0] if urls else None
        else:
            return None
    else:
        oFilePicker.MultiSelectionMode = True
        if oFilePicker.execute() == ExecutableDialogResults.OK:
            return oFilePicker.SelectedFiles
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
        x = self._centered(self._dialog_rectangle.w, self._bar_dimensions.width)
        _set_rectangle(self._oBarModel,
                       Rectangle(x, MARGIN, self._bar_dimensions.width,
                                 self._bar_dimensions.height))
        self._oBarModel.ProgressValueMin = self._bar_progress.min
        self._oBarModel.ProgressValueMax = self._bar_progress.max
        self._oBarModel.ProgressValue = self._bar_progress.min
        y = self._bar_dimensions.height + MARGIN * 2
        w = self._dialog_rectangle.w - MARGIN * 2
        h = self._dialog_rectangle.h - self._bar_dimensions.height - MARGIN * 2
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
        self._oBar.Value = 0
        self._oText = oText
        self.response = None

    def progress(self, n: int = 1):
        self._oBar.Value = self._oBar.Value + n

    def message(self, text: str):
        self._oText.Text = text


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
