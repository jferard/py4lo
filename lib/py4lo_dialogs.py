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
# mypy: disable-error-code="import-untyped,import-not-found"
from collections import namedtuple
from enum import Enum
from threading import Thread
from typing import Any, Callable, Optional, List, Union, NamedTuple

from py4lo_helper import (create_uno_service_ctxt, provider,
                          create_uno_service, create_uno_struct)
from py4lo_typing import UnoControlModel, UnoControl, StrPath

try:
    # noinspection PyUnresolvedReferences
    import uno


    class MessageBoxType:
        # noinspection PyUnresolvedReferences
        from com.sun.star.awt.MessageBoxType import (
            MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX, )


    class MessageBoxButtons:
        # noinspection PyUnresolvedReferences
        from com.sun.star.awt.MessageBoxButtons import (
            BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO,
            BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL,
            BUTTONS_ABORT_IGNORE_RETRY, DEFAULT_BUTTON_OK,
            DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, DEFAULT_BUTTON_YES,
            DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE,
        )


    class FontWeight:
        # noinspection PyUnresolvedReferences
        from com.sun.star.awt.FontWeight import (BOLD, )


    class ExecutableDialogResults:
        # noinspection PyUnresolvedReferences
        from com.sun.star.ui.dialogs.ExecutableDialogResults import (
            OK, CANCEL)


    class PushButtonType:
        # noinspection PyUnresolvedReferences
        from com.sun.star.awt.PushButtonType import (OK, CANCEL)

except (ModuleNotFoundError, ImportError):
    from mock_constants import (  # type: ignore[assignment]
        ExecutableDialogResults,  # pyright: ignore[reportGeneralTypeIssues]
        MessageBoxButtons,  # pyright: ignore[reportGeneralTypeIssues]
        MessageBoxType,  # pyright: ignore[reportGeneralTypeIssues]
        PushButtonType,  # pyright: ignore[reportGeneralTypeIssues]
        uno,  # pyright: ignore[reportGeneralTypeIssues]
    )


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
        oWidgetModel: UnoControlModel, x: int, y: int,
        width: int, height: int):
    """Place a widget on the widget model"""
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
    oTextControl = create_uno_service(Control.FixedText)
    oTextModel = oDialogModel.createInstance(ControlModel.FixedText)
    oTextModel.Label = text
    oTextControl.setModel(oTextModel)
    min_size = oTextControl.MinimumSize
    # Why 0.5 ? I don't know
    return Size(min_size.Width * 0.5, min_size.Height * 0.5)


def message_box(msg_title: str, msg_text: str,
                msg_type=MessageBoxType.MESSAGEBOX,
                msg_buttons=MessageBoxButtons.BUTTONS_OK,
                parent_win=None) -> int:
    """Create a message box"""
    # from https://forum.openoffice.org/fr/forum/viewtopic.php?f=15&t=47603#
    # (thanks Bernard !)
    toolkit = create_uno_service_ctxt("com.sun.star.awt.Toolkit")
    if parent_win is None:
        parent_win = provider.parent_win
    mb = toolkit.createMessageBox(parent_win, msg_type, msg_buttons, msg_title,
                                  msg_text)
    return mb.execute()


class InputBox:
    # see https://wiki.documentfoundation.org/Macros
    # /General/IO_to_Screen#Using_Application_Programming_Interface_(API)
    """TODO: choose input type"""

    def __init__(
            self, width: int, height: int, hori_margin: int, vert_margin: int,
            button_width: int, button_height: int,
            hori_sep: int, vert_sep: int, label_width: int, label_height: int,
            edit_height: int):
        self.width = width
        self.height = height
        self.hori_margin = hori_margin
        self.vert_margin = vert_margin
        self.button_width = button_width
        self.button_height = button_height
        self.hori_sep = hori_sep
        self.vert_sep = vert_sep
        self.label_width = label_width
        self.label_height = label_height
        self.edit_height = edit_height

    def input(self, msg_title: str, msg_text: str, msg_default: str = "",
              parent_win=None, x: Optional[int] = None,
              y: Optional[int] = None) -> Optional[str]:
        if parent_win is None:
            parent_win = provider.parent_win
        toolkit = create_uno_service_ctxt("com.sun.star.awt.Toolkit")

        if x is None or y is None:
            ps = parent_win.PosSize
            if x is None:
                x = (ps.Width - self.width) / 2
            if y is None:
                y = (ps.Height - self.height) / 2

        oDialogModel = create_uno_service(ControlModel.Dialog)
        oDialogModel.Title = msg_title
        place_widget(oDialogModel, x, y, self.width, self.height)

        self._create_label_model(oDialogModel, "label", msg_text)
        oEditModel = self._create_edit_model(
            oDialogModel, "edit", msg_default)
        self._create_cancel_model(oDialogModel, "btn_cancel")
        self._create_ok_model(oDialogModel, "btn_ok")

        dialog = create_uno_service(Control.Dialog)
        dialog.setModel(oDialogModel)

        dialog.createPeer(toolkit, parent_win)

        oEditControl = dialog.getControl("edit")
        oEditControl.setSelection(
            create_uno_struct("com.sun.star.awt.Selection", Min=0,
                              Max=len(msg_default)))
        oEditControl.setFocus()

        if dialog.execute() == ExecutableDialogResults.CANCEL:
            return None

        ret = oEditModel.Text
        dialog.dispose()
        return ret

    def _create_label_model(self, oDialogModel: UnoControlModel, name: str,
                            msg_text: str) -> UnoControlModel:
        oLabelModel = oDialogModel.createInstance(ControlModel.FixedText)
        place_widget(oLabelModel, self.hori_margin, self.vert_margin,
                     self.label_width, self.label_height)
        oLabelModel.Label = msg_text
        oLabelModel.NoLabel = True
        oDialogModel.insertByName(name, oLabelModel)
        return oLabelModel

    def _create_edit_model(self, oDialogModel, name, msg_default):
        oEditModel = oDialogModel.createInstance(
            "com.sun.star.awt.UnoControlEditModel")
        edit_width = self.width - self.hori_margin * 2
        place_widget(oEditModel,
                     self.hori_margin,
                     self.vert_margin + self.label_height + self.vert_sep,
                     edit_width, self.edit_height)
        oEditModel.Text = msg_default
        oDialogModel.insertByName(name, oEditModel)
        return oEditModel

    def _create_cancel_model(self, oDialogModel: UnoControlModel,
                             name: str) -> UnoControlModel:
        oCancelModel = oDialogModel.createInstance(ControlModel.Button)
        place_widget(oCancelModel,
                     self.width - (
                             self.hori_margin + self.button_width
                             + self.hori_sep + self.button_width),
                     (self.vert_margin + self.label_height + self.vert_sep +
                      self.edit_height + self.vert_sep),
                     self.button_width, self.button_height)
        oCancelModel.PushButtonType = PushButtonType.CANCEL
        oCancelModel.DefaultButton = False
        oDialogModel.insertByName(name, oCancelModel)
        return oCancelModel

    def _create_ok_model(self, oDialogModel: UnoControlModel,
                         name: str) -> UnoControlModel:
        oOkModel = oDialogModel.createInstance(ControlModel.Button)
        place_widget(oOkModel,
                     self.width - (self.hori_margin + self.button_width),
                     (self.vert_margin + self.label_height
                      + self.vert_sep + self.edit_height + self.vert_sep),
                     self.button_width, self.button_height)
        oOkModel.PushButtonType = PushButtonType.OK
        oOkModel.DefaultButton = True
        oDialogModel.insertByName(name, oOkModel)
        return oOkModel


class InputBoxBuilder:
    # see https://wiki.documentfoundation.org/Macros/General/
    # IO_to_Screen#Using_Application_Programming_Interface_(API)
    def __init__(self):
        self.width = 300
        self.height = None
        self.hori_margin = 8
        self.vert_margin = 8
        self.label_height = 10
        self.edit_height = 20
        self.button_width = 60
        self.button_height = 18
        self.hori_sep = 4
        self.vert_sep = 4

    def _set_sep(self, value: int):
        self.hori_sep = value
        self.vert_sep = value

    def _set_margin(self, value: int):
        self.hori_margin = value
        self.vert_margin = value

    sep = property(fset=_set_sep)
    margin = property(fset=_set_margin)

    def build(self) -> InputBox:
        if self.height is None:
            self.height = (
                    self.vert_margin + self.label_height + self.vert_sep
                    + self.edit_height + self.vert_sep + self.button_height
                    + self.vert_margin)
        else:
            self.label_height = self.height - (
                    self.vert_margin + self.vert_sep + self.edit_height
                    + self.vert_sep + self.button_height + self.vert_margin)

        label_width = self.width - (
                self.button_width + self.hori_sep + self.hori_margin * 2)
        return InputBox(
            self.width, self.height, self.hori_margin, self.vert_margin,
            self.button_width, self.button_height,
            self.hori_sep, self.vert_sep, label_width, self.label_height,
            self.edit_height)


def input_box(msg_title: str, msg_text: str, msg_default="", parent_win=None,
              x: Optional[int] = None,
              y: Optional[int] = None) -> str:
    """Create an input box"""
    return InputBoxBuilder().build().input(msg_title, msg_text, msg_default,
                                           parent_win, x, y)


FileFilter = NamedTuple("FileFilter", [("title", str), ("filter", str)])


def file_dialog(title: str, filters: Optional[List[FileFilter]] = None,
                display_dir: StrPath = "",
                single: bool = True) -> Union[Optional[str], List[str]]:
    """
    Open a file dialog
    @return: if single, url or None, else a list of urls
    """
    oFilePicker = create_uno_service("com.sun.star.ui.dialogs.FilePicker")
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
    oFolder = create_uno_service("com.sun.star.ui.dialogs.FolderPicker")
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
        self._oDialogModel = create_uno_service(ControlModel.Dialog)
        self._oBarModel = self._oDialogModel.createInstance(
            ControlModel.ProgressBar)
        self._oTextModel = self._oDialogModel.createInstance(
            ControlModel.FixedText)
        self._oDialog = create_uno_service(Control.Dialog)
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
        x = self._centered(
            self._dialog_rectangle.w, self._bar_dimensions.width)
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
                                self._bar_progress.min, self._bar_progress.max,
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


class VoidProgressHandler:
    """
    A progress handler
    """

    def progress(self, n: int):
        """
        Update the progress
        @param n: number of steps since the last progress
        """
        pass

    def set(self, i: int):
        """
        Set the progress
        @param i: total number of steps since the beginning
        """
        pass

    def reset(self):
        """
        Reset the progress
        """
        pass

    def message(self, text: str):
        """
        Create a message
        @param text: the text
        """
        pass


class ProgressHandler(VoidProgressHandler):
    def __init__(self, oBar: UnoControlModel, bar_progress_min: int,
                 bar_progress_max: int, oText: UnoControlModel):
        self._oBar = oBar
        self._oBar.Value = bar_progress_min
        self._bar_progress_min = bar_progress_min
        self._bar_progress_max = bar_progress_max
        self._oText = oText
        self.response = None

    def progress(self, n: int = 1):
        self._oBar.Value = self._oBar.Value + n
        if self._oBar.Value > self._bar_progress_max:
            self._oBar.Value = self._bar_progress_max

    def set(self, i: int):
        if i > self._bar_progress_max:
            self._oBar.Value = self._bar_progress_max
        else:
            self._oBar.Value = i

    def reset(self):
        self._oBar.Value = self._bar_progress_min

    def message(self, text: str):
        self._oText.Text = text


class ProgressExecutor:
    def __init__(self, oDialog: UnoControl, autoclose: bool,
                 oBar: UnoControlModel, bar_progress_min: int,
                 bar_progress_max: int,
                 oText: UnoControlModel):
        self._oDialog = oDialog
        self._autoclose = autoclose
        self._progress_handler = ProgressHandler(oBar, bar_progress_min,
                                                 bar_progress_max, oText)

    def execute(self, func: Callable[[ProgressHandler], None]):
        """
        Execute the function with a progress bar
        @param func: a function that takes a `ProgressDialog` object.
        """
        toolkit = create_uno_service("com.sun.star.awt.Toolkit")

        def aux():
            self._oDialog.setVisible(True)
            self._oDialog.createPeer(toolkit, None)
            func(self._progress_handler)
            if self._autoclose:
                # free all resources as soon as the function is executed
                self._oDialog.dispose()
            else:
                # wait for the user to close the window
                self._oDialog.execute()

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
        self._oDialogModel = create_uno_service(ControlModel.Dialog)
        self._oDialogModel.Closeable = True
        self._oDialog = create_uno_service(Control.Dialog)
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
        _set_rectangle(self._oTextModel, Rectangle(
            MARGIN, MARGIN,
            self._console_rectangle.w - MARGIN * 2,
            self._console_rectangle.h - MARGIN * 2))
        return ConsoleExecutor(self._oDialog, self._autoclose,
                               self._oDialog.getControl("text"))


class VoidConsoleHandler:
    """
    A console handler
    """

    def message(self, text: str):
        """
        A message
        @param text: the text of the message
        """
        pass


class ConsoleHandler(VoidConsoleHandler):
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
        toolkit = create_uno_service("com.sun.star.awt.Toolkit")

        def aux():
            self._oDialog.setVisible(True)
            self._oDialog.createPeer(toolkit, None)
            func(self._console_handler)
            if self._autoclose:
                # free all resources as soon as the function is executed
                self._oDialog.dispose()
            else:
                # wait for the user to close the window
                self._oDialog.execute()

        t = Thread(target=aux)
        t.start()

    @property
    def response(self):
        return self._console_handler.response
