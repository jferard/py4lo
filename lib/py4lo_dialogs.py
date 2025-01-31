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
"""
The module py4lo_dialogs gives functions to create a LibreOffice dialog from
scratch.

Here's a full example:

```
oDialogModel = cast(UnoControlModel,
                    create_uno_service(ControlModel.Dialog))
oDialogModel.Title = "The title"
place_widget(oDialogModel, 100, 100, 400, 300)

oLabelModel = cast(UnoControlModel,
                   oDialogModel.createInstance(ControlModel.FixedText))
place_widget(oLabelModel, 10, 10, 390, 20)
oLabelModel.Label = "A text"
oDialogModel.insertByName("label", oLabelModel)

oDialog = create_uno_service(Control.Dialog)
oDialog.setModel(oDialogModel)

oOkModel = cast(UnoControlModel,
                oDialogModel.createInstance(ControlModel.Button))
place_widget(oOkModel, 100, 350, 100, 20)
oOkModel.PushButtonType = PushButtonType.OK
oOkModel.DefaultButton = True
oDialogModel.insertByName("button", oOkModel)

oToolkit = get_toolkit()
oDialog.createPeer(oToolkit, None)

if oDialog.execute() == ExecutableDialogResults.OK:
    ...

oDialog.dispose()
```

To load an XML dialog (like in Basic), see: `_ObjectProvider.get_dialog`
from py4lo_helper.
"""
import functools
import logging
# mypy: disable-error-code="import-untyped,import-not-found"

from collections import namedtuple
from enum import Enum
from threading import Thread
from typing import Any, Callable, Optional, List, Union, NamedTuple, cast

from py4lo_helper import (
    get_provider, create_uno_service, create_uno_struct)
from py4lo_typing import (
    UnoControlModel, UnoControl, StrPath, lazy, UnoService)

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


    class MessageBoxResults:
        # noinspection PyUnresolvedReferences
        from com.sun.star.awt.MessageBoxResults import (
            CANCEL, OK, YES, NO, RETRY, IGNORE
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

except ImportError:
    from _mock_constants import (  # type: ignore[assignment]
        ExecutableDialogResults,  # pyright: ignore[reportGeneralTypeIssues]
        MessageBoxButtons,  # pyright: ignore[reportGeneralTypeIssues]
        MessageBoxType,  # pyright: ignore[reportGeneralTypeIssues]
        PushButtonType,  # pyright: ignore[reportGeneralTypeIssues]
        MessageBoxResults,  # pyright: ignore[reportGeneralTypeIssues]  # noqa: F401
    )
    from _mock_objects import (  # type: ignore[assignment]
        uno,  # pyright: ignore[reportGeneralTypeIssues]
    )


class ControlModel(str, Enum):
    """List of UNO control models names"""
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
    """List of UNO control names"""
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

_oToolkit = lazy(UnoService)


def get_toolkit() -> UnoService:
    """
    @return: the com.sun.star.awt.Toolkit instance
    """
    global _oToolkit
    if _oToolkit is None:
        _oToolkit = create_uno_service("com.sun.star.awt.Toolkit")
    return _oToolkit

###
# Common functions
###

def place_widget(
        oWidgetModel: UnoControlModel, x: int, y: int,
        width: int, height: int):
    """
    Place a widget on the widget model (see
    com.sun.star.awt.UnoControlDialogElement)

    @param oWidgetModel: the model of the widget to place
    @param x: the x position
    @param y: the y position
    @param width: the width of the widget
    @param height: the height of the widget
    """
    oWidgetModel.PositionX = x
    oWidgetModel.PositionY = y
    oWidgetModel.Width = width
    oWidgetModel.Height = height


Size = namedtuple("Size", ["width", "height"])


def get_text_size(oDialogModel: UnoControlModel, text: str) -> Size:
    """
    Return the size of a box containing the given text on a dialog
    model.

    @param oDialogModel: the model
    @param text: the text
    @return: the text size (see com.sun.star.awt.Size)
    """
    oTextControl = create_uno_service(Control.FixedText)
    oTextModel = cast(UnoControlModel,
                      oDialogModel.createInstance(ControlModel.FixedText))
    oTextModel.Label = text
    oTextControl.setModel(oTextModel)
    min_size = oTextControl.MinimumSize
    # Why 0.5? I don't know
    return Size(min_size.Width * 0.5, min_size.Height * 0.5)


def message_box(msg_title: str, msg_text: str,
                msg_type=MessageBoxType.MESSAGEBOX,
                msg_buttons=MessageBoxButtons.BUTTONS_OK,
                parent_win=None) -> MessageBoxResults:
    """
    Create a message box

    @param msg_title: the title of the message box
    @param msg_text: the text
    @param msg_type: the type (see com.sun.star.awt.MessageBoxType)
    @param msg_buttons: the buttons (see com.sun.star.awt.MessageBoxButtons)
    @param parent_win: the optional parent window
    @return: the result (see com.sun.star.awt.MessageBoxResults)
    """
    # from https://forum.openoffice.org/fr/forum/viewtopic.php?f=15&t=47603#
    # (thanks Bernard !)
    oToolkit = get_toolkit()
    if parent_win is None:
        parent_win = get_provider().parent_win
    mb = oToolkit.createMessageBox(parent_win, msg_type, msg_buttons, msg_title,
                                  msg_text)
    return mb.execute()


class InputBox:
    """
    A very ugly input box factory.
    Use `InputBoxBuilder`.

    Example:
    ```
    builder = InputBoxBuilder()
    builder.width = 200
    factory = builder.build()
    factory.input("Game", "Guess the number (1-1000)")
    """

    # see https://wiki.documentfoundation.org/Macros/General/IO_to_Screen
    def __init__(
            self, width: int, height: int, hori_margin: int, vert_margin: int,
            button_width: int, button_height: int,
            hori_sep: int, vert_sep: int, label_width: int, label_height: int,
            edit_height: int):
        """
        @param width: the width of the box
        @param height: the height of the box
        @param hori_margin: the left and right margins
        @param vert_margin: the top and bottom margins
        @param button_width: the width of the OK / Cancel buttons
        @param button_height: the height of the OK / Cancel buttons
        @param hori_sep: the space between the buttons
        @param vert_sep: the space between the label and the text box
        @param label_width: the width of the label
        @param label_height: the height of the label
        @param edit_height: the height of the edit box
        """
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
        """
        Execute an input box.

        @param msg_title: the title of the input box
        @param msg_text: the text
        @param msg_default: the default text inside the box
        @param parent_win: the optional parent window
        @param x: the x position of the box
        @param y: the y position of the box
        @return: the value inside the input box
        """
        if parent_win is None:
            parent_win = get_provider().parent_win
        oToolkit = get_toolkit()

        if x is None or y is None:
            ps = parent_win.PosSize
            if x is None:
                x = int((ps.Width - self.width) / 2)
            if y is None:
                y = int((ps.Height - self.height) / 2)

        oDialogModel = cast(UnoControlModel,
                            create_uno_service(ControlModel.Dialog))
        oDialogModel.Title = msg_title
        place_widget(oDialogModel, x, y, self.width, self.height)

        self._create_label_model(oDialogModel, "label", msg_text)
        _oEditModel = self._create_edit_model(
            oDialogModel, "edit", msg_default)
        self._create_cancel_model(oDialogModel, "btn_cancel")
        self._create_ok_model(oDialogModel, "btn_ok")

        oDialog = create_uno_service(Control.Dialog)
        oDialog.setModel(oDialogModel)

        oDialog.createPeer(oToolkit, parent_win)

        oEditControl = oDialog.getControl("edit")
        oEditControl.setSelection(
            create_uno_struct("com.sun.star.awt.Selection", Min=0,
                              Max=len(msg_default)))
        oEditControl.setFocus()

        if oDialog.execute() == ExecutableDialogResults.CANCEL:
            return None

        ret = oEditControl.Text
        oDialog.dispose()
        return ret

    def _create_label_model(self, oDialogModel: UnoControlModel, name: str,
                            msg_text: str) -> UnoControlModel:
        oLabelModel = cast(UnoControlModel,
                           oDialogModel.createInstance(ControlModel.FixedText))
        place_widget(oLabelModel, self.hori_margin, self.vert_margin,
                     self.label_width, self.label_height)
        oLabelModel.Label = msg_text
        oLabelModel.NoLabel = True
        oDialogModel.insertByName(name, oLabelModel)
        return oLabelModel

    def _create_edit_model(
            self, oDialogModel: UnoControlModel, name: str, msg_default: str
    ) -> UnoControlModel:
        oEditModel = cast(UnoControlModel, oDialogModel.createInstance(
            "com.sun.star.awt.UnoControlEditModel"))
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
        oCancelModel = cast(UnoControlModel,
                            oDialogModel.createInstance(ControlModel.Button))
        x = self.width - (
                self.hori_margin + self.button_width +
                self.hori_sep + self.button_width
        )
        y = (
                self.vert_margin + self.label_height +
                self.vert_sep + self.edit_height + self.vert_sep
        )
        place_widget(oCancelModel,
                     x, y, self.button_width, self.button_height)
        oCancelModel.PushButtonType = PushButtonType.CANCEL
        oCancelModel.DefaultButton = False
        oDialogModel.insertByName(name, oCancelModel)
        return oCancelModel

    def _create_ok_model(self, oDialogModel: UnoControlModel,
                         name: str) -> UnoControlModel:
        oOkModel = cast(UnoControlModel,
                        oDialogModel.createInstance(ControlModel.Button))
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
    """
    A very ugly input box factory builder.

    Example:
    ```
    builder = InputBoxBuilder()
    builder.width = 200
    factory = builder.build()
    factory.input("Game", "Guess the number (1-1000)")
    ```
    """

    # see https://wiki.documentfoundation.org/Macros/General/IO_to_Screen
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
        """
        @return: the input box factory
        """
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
    """
    Execute an input box.

    @param msg_title: the title of the input box
    @param msg_text: the text
    @param msg_default: the default text inside the box
    @param parent_win: the optional parent window
    @param x: the x position of the box
    @param y: the y position of the box
    @return: the value inside the input box
    """
    return InputBoxBuilder().build().input(msg_title, msg_text, msg_default,
                                           parent_win, x, y)


FileFilter = NamedTuple("FileFilter", [("title", str), ("filter", str)])


def file_dialog(title: str, filters: Optional[List[FileFilter]] = None,
                display_dir: StrPath = "",
                single: bool = True) -> Union[Optional[str], List[str]]:
    """
    Open a file dialog.

    Example:
    ```
    urls = file_dialog(
        "Choose csv files", [FileFilter("CSV", "*.csv")], single=False)
    ```

    @param title: the title of the dialog
    @param filters: the filter
    @param display_dir: the base directory of the dialog
    @param single: if True, select one files, otherwise allows multiple selction.
    @return: if single is True, url or None, else a list of urls
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
    Open a folder/directory dialog.

    Example:
    ```
    urls = folder_dialog("Choose csv files folder")
    ```

    @param title: the title of the dialog
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
Rectangle = NamedTuple("Rectangle",
                       [("x", int), ("y", int), ("w", int), ("h", int)])
Progress = NamedTuple("Progress",
                      [("min", int), ("max", int)])


def _set_rectangle(oWidgetModel: UnoControlModel, rectangle: Rectangle):
    """
    Place a widget on the widget model (see
    com.sun.star.awt.UnoControlDialogElement)

    @param oWidgetModel: the model of the widget to place
    @param rectangle: the position and size of the widget
    """
    oWidgetModel.PositionX = rectangle.x
    oWidgetModel.PositionY = rectangle.y
    oWidgetModel.Width = rectangle.w
    oWidgetModel.Height = rectangle.h


class ProgressActionStopped(Exception):
    """
    Exception thrown by a handler or a function to stop the progress action.
    """
    pass


class ProgressHandler:
    """
    Any object can be a handler.
    """


# ProgressExecutor
class VoidProgressHandler(ProgressHandler):
    """
    A VoidProgressHandler. This is the base class for StandardProgressHandler.

    A progress handler allows to trigger a progress, to set a progress value
    or to set a progress text.

    Typically, it is passed to a function as the only parameter:

    ```
    progress_handler = VoidProgressHandler() # or another handler
    def myfunc(progress_handler: VoidProgressHandler):
        do_something()
        progress_handler.progress(10)
        progress_handler.message("Something done")
        ...
    ```

    The function may set the `response` attribute of the progress_handler to
    return a value.
    """

    def __init__(self):
        self.response = None

    def progress(self, n: int):
        """
        Update the progress value
        @param n: number of steps since the last progress
        """
        pass

    def set(self, i: int):
        """
        Set the progress value
        @param i: total number of steps since the beginning
        """
        pass

    def reset(self):
        """
        Reset the progress value
        """
        pass

    def message(self, text: str):
        """
        Set a progress message
        @param text: the text
        """
        pass


class StandardProgressHandler(VoidProgressHandler):
    """
    A StandardProgressHandler will send progress value updates to a progress bar
    and the progress message to a text box.

    Typically, it is passed to a function as the only parameter:

    ```
    progress_handler = StandardProgressHandler(...) # or another handler
    def myfunc(progress_handler: VoidProgressHandler):
        do_something()
        progress_handler.progress(10)
        progress_handler.message("Something done")
        ...
    ```

    The function may set the `response` attribute of the progress_handler to
    return a value.
    """

    def __init__(self, oBar: UnoControl, bar_progress_min: int,
                 bar_progress_max: int, oText: UnoControl):
        """
        @param oBar: the progress bar (see: com.sun.star.awt.UnoControlProgressBar)
        @param bar_progress_min: the minimum progress value
        @param bar_progress_max: the maximum progress value
        @param oText: the progress text box (see: com.sun.star.awt.UnoControlFixedTextModel)
        """
        VoidProgressHandler.__init__(self)
        self.oBar = oBar
        self.oBar.Value = bar_progress_min
        self._bar_progress_min = bar_progress_min
        self._bar_progress_max = bar_progress_max
        self._oText = oText
        self.response = None

    def progress(self, n: int = 1):
        self.oBar.Value = self.oBar.Value + n
        if self.oBar.Value > self._bar_progress_max:
            self.oBar.Value = self._bar_progress_max

    def set(self, i: int):
        if i > self._bar_progress_max:
            self.oBar.Value = self._bar_progress_max
        else:
            self.oBar.Value = i

    def reset(self):
        self.oBar.Value = self._bar_progress_min

    def message(self, text: str):
        self._oText.Text = text


class ProgressExecutor:
    """
    A ProgressExecutor takes a dialog that contains a progress bar
    and a progress text box. It will then execute a function.

    See ProgressExecutorBuilder for a convenient way to create this object.

    ```
    def myfunc(progress_handler: VoidProgressHandler):
        do_something()
        progress_handler.progress(10)
        progress_handler.message("Something done")
        ...
        progress_handler.response = "OK"

    executor = ProgressExecutor.create(...)
    executor.execute(func)
    x = executor.response
    ```
    """
    _logger = logging.getLogger(__name__)

    @staticmethod
    def create(oDialog: UnoControl, autoclose: bool,
               oBar: UnoControl, bar_progress_min: int,
               bar_progress_max: int, oText: UnoControl) -> "ProgressExecutor":
        """
        @param oDialog: the dialog containing the bar and the text box
        @param autoclose: if True, close the dialog when the function is executed
        @param oBar: the progress bar (see: com.sun.star.awt.UnoControlProgressBar)
        @param bar_progress_min: the minimum progress value
        @param bar_progress_max: the maximum progress value
        @param oText: the progress text box (see: com.sun.star.awt.UnoControlFixedTextModel)
        """
        progress_handler = StandardProgressHandler(
            oBar, bar_progress_min, bar_progress_max, oText)
        return ProgressExecutor(oDialog, autoclose, progress_handler)

    def __init__(self, oDialog: UnoControl, autoclose: bool,
                 progress_handler: ProgressHandler):
        """
        @param oDialog: the dialog containing the bar and the text box
        @param autoclose: if True, close the dialog when maximum is reached
        @param progress_handler: the progress handler
        """
        self._oDialog = oDialog
        self._autoclose = autoclose
        self._progress_handler = progress_handler

    def execute(self, func: Callable[[ProgressHandler], None]):
        """
        Execute the function in a thread and reverberate StandardProgressHandler
        commands to the dialog.

        @param func: a function that takes a `StandardProgressHandler` object. The
        function may set the `response` attribute of the progress_handler to
        return a value.
        """
        oToolkit = get_toolkit()

        def aux():
            self._oDialog.setVisible(True)
            self._oDialog.createPeer(oToolkit, None)
            try:
                func(self._progress_handler)
            except ProgressActionStopped:
                pass
            except Exception:
                self._logger.exception("Progress exception!")
            finally:
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
        Get the response from the handler.

        @return: the response
        """
        return self._progress_handler.response


class ProgressDialogBuilder:
    """
    A dialog builder for ProgessExecutor.

    Example:
    ```
    builder = ProgressDialogBuilder()
    oDialog = builder.title("See the progress").build()
    ```
    """

    def __init__(self):
        self._oDialogModel = cast(
            UnoControlModel, create_uno_service(ControlModel.Dialog))
        self._oBarModel = cast(
            UnoControlModel,
            self._oDialogModel.createInstance(ControlModel.ProgressBar))
        self._oTextModel = cast(UnoControlModel,
                                self._oDialogModel.createInstance(
                                    ControlModel.FixedText))
        self._oDialog = cast(UnoControl,
                             create_uno_service(Control.Dialog))
        self.title("Please wait...")
        self._dialog_rectangle = Rectangle(150, 150, 150, 40)
        self._bar_dimensions = Size(140, 12)
        self._bar_progress = Progress(0, 100)
        self._message = None

    def build(self) -> UnoControl:
        """
        @return: the executor
        """
        self._oDialogModel.insertByName("bar", self._oBarModel)
        self._oDialogModel.insertByName("text", self._oTextModel)
        _set_rectangle(self._oDialogModel, self._dialog_rectangle)
        self._oDialog.setModel(self._oDialogModel)
        x = self._centered(
            self._dialog_rectangle.w, self._bar_dimensions.width)
        bar_rectangle = Rectangle(
            x, MARGIN, self._bar_dimensions.width, self._bar_dimensions.height)
        _set_rectangle(self._oBarModel, bar_rectangle)
        self._oBarModel.ProgressValueMin = self._bar_progress.min
        self._oBarModel.ProgressValueMax = self._bar_progress.max
        self._oBarModel.ProgressValue = self._bar_progress.min
        y = self._bar_dimensions.height + MARGIN * 2
        w = self._dialog_rectangle.w - MARGIN * 2
        h = self._dialog_rectangle.h - self._bar_dimensions.height - MARGIN * 2
        _set_rectangle(self._oTextModel, Rectangle(MARGIN, y, w, h))
        if self._message is not None:
            self._oTextModel.Label = self._message
        return self._oDialog

    def _centered(self, outer_w: int, inner_w: int) -> int:
        if outer_w <= inner_w:
            return 0
        else:
            return (outer_w - inner_w) // 2

    def title(self, title: str) -> "ProgressDialogBuilder":
        """
        Set the title
        @param title: the title

        @return: self for fluent style
        """
        self._oDialogModel.Title = title
        return self

    def dialog_rectangle(self, x: int, y: int, w: int,
                         h: int) -> "ProgressDialogBuilder":
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

    def bar_dimensions(self, w: int, h: int) -> "ProgressDialogBuilder":
        """
        Set the dialog rectangle

        @param w: width
        @param h: height
        @return: self for fluent style
        """
        self._bar_dimensions = Size(w, h)
        return self

    def bar_progress(self, progress_min: int,
                     progress_max: int) -> "ProgressDialogBuilder":
        """
        Set the dialog rectangle

        @param progress_min: min
        @param progress_max: max
        @return: self for fluent style
        """
        self._bar_progress = Progress(progress_min, progress_max)
        return self

    def message(self, message: str) -> "ProgressDialogBuilder":
        """
        Set the message
        @param message: the message
        @return: self for fluent style
        """
        self._message = message
        return self


class ProgressExecutorBuilder:
    """
    A ProgessExecutorBuilder.

    Example:
    ```
    builder = ProgressExecutorBuilder()
    executor = builder.title("See the progress").autoclose(True).build()
    ```
    """

    def __init__(self):
        self._dialog_builder = ProgressDialogBuilder()
        self._progress_handler = lazy(ProgressHandler)
        self._autoclose = True

    def build(self) -> ProgressExecutor:
        """
        @return: the executor
        """
        oDialog = self._dialog_builder.build()
        if self._progress_handler is None:
            oBar = oDialog.getControl("bar")
            oText = oDialog.getControl("text")
            bar_progress_min = oBar.Model.ProgressValueMin
            bar_progress_max = oBar.Model.ProgressValueMax
            progress_handler = StandardProgressHandler(
                oBar, bar_progress_min, bar_progress_max,
                oText
            )
        else:
            progress_handler = self._progress_handler

        return ProgressExecutor(oDialog, self._autoclose, progress_handler)

    def handler(self, handler: ProgressHandler) -> "ProgressExecutorBuilder":
        """
        Sets a specific handler for this executor
        @param handler: the handler
        @return: self for fluent style
        """
        self._progress_handler = handler
        return self

    def title(self, title: str) -> "ProgressExecutorBuilder":
        """See ProgressDialogBuilder.title"""
        self._dialog_builder.title(title)
        return self

    def autoclose(self, b: bool) -> "ProgressExecutorBuilder":
        """
        Set the autoclose value : if False, don't close
        @param b: the value of autoclose
        @return: self for fluent style
        """
        self._autoclose = b
        return self

    def dialog_rectangle(self, x: int, y: int, w: int,
                         h: int) -> "ProgressExecutorBuilder":
        """See ProgressDialogBuilder.dialog_rectangle"""
        self._dialog_builder.dialog_rectangle(x, y, w, h)
        return self

    def bar_dimensions(self, w, h):
        """See ProgressDialogBuilder.bar_dimensions"""
        self._dialog_builder.bar_dimensions(w, h)
        return self

    def bar_progress(self, progress_min: int,
                     progress_max: int) -> "ProgressExecutorBuilder":
        """See ProgressDialogBuilder.bar_progress"""
        self._dialog_builder.bar_progress(progress_min, progress_max)
        return self

    def message(self, message: str) -> "ProgressExecutorBuilder":
        """See ProgressDialogBuilder.message"""
        self._dialog_builder.message(message)
        return self


# ConsoleExecutor

class ConsoleHandler:
    """
    Base class for console handlers
    """
    pass


class VoidConsoleHandler(ConsoleHandler):
    """
    A VoidConsoleHandler. This is the base class for StandardConsoleHandler objects.

    Typically, it is passed to a function as the only parameter:

    ```
    progress_handler = VoidConsoleHandler()
    def myfunc(console_handler: VoidConsoleHandler):
        do_something()
        console_handler.message("Something done")
        ...
    ```

    The function may set the `response` attribute of the coonsole_handler to
    return a value.
    """

    def __init__(self):
        self.response = None

    def message(self, text: str):
        """
        A message
        @param text: the text of the message
        """
        pass


class StandardConsoleHandler(VoidConsoleHandler):
    """
    A StandardConsoleHandler will send the progress message to a text box.

    Typically, it is passed to a function as the only parameter:

    ```
    progress_handler = StandardConsoleHandler(...)
    def myfunc(console_handler: StandardConsoleHandler):
        do_something()
        console_handler.message("Something done")
        ...
    ```

    The function may set the `response` attribute of the coonsole_handler to
    return a value.
    """

    def __init__(self, oText: UnoControl):
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
    """
    A ConsoleExecutor takes a dialog that contains a progress text box.
    It will then execute a function.

    See ConsoleExecutorBuilder for a convenient way to create this object.

    ```
    def myfunc(console_handler: VoidProgressHandler):
        do_something()
        console_handler.message("Something done")
        ...
        console_handler.response = "OK"

    executor = ConsoleExecutor.create(...)
    executor.execute(func)
    x = executor.response
    ```
    """
    def __init__(self, oDialog: UnoControl, autoclose: bool,
                 console_handler: VoidConsoleHandler):
        """
        @param oDialog: the dialog containing the text box
        @param autoclose: if True, close the dialog when the function is executed
        @param console_handler: the StandardConsoleHandler
        """
        self._oDialog = oDialog
        self._autoclose = autoclose
        self._console_handler = console_handler

    def execute(self, func: Callable[[VoidConsoleHandler], None]):
        """
        Execute the function in a thread and reverberate StandardConsoleHandler
        commands to the dialog.

        @param func: a function that takes a `StandardConsoleHandler` object. The
        function may set the `response` attribute of the progress_handler to
        return a value.
        """
        oToolkit = get_toolkit()

        def aux():
            self._oDialog.setVisible(True)
            self._oDialog.createPeer(oToolkit, None)
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
        """
        Get the response from the handler.

        @return: the response
        """
        return self._console_handler.response


class ConsoleDialogBuilder:
    """
    A ConsoleExecutorBuilder

    Example:
    ```
    builder = ConsoleExecutorBuilder()
    executor = builder.title("See the messages").autoclose(True).build()
    ```
    """

    def __init__(self):
        self._oDialogModel = cast(
            UnoControlModel, create_uno_service(ControlModel.Dialog))
        self._oDialogModel.Closeable = True
        self._oDialog = cast(
            UnoControl, create_uno_service(Control.Dialog))
        self._oTextModel = cast(
            UnoControlModel,
            self._oDialogModel.createInstance(ControlModel.Edit))
        self._oTextModel.ReadOnly = True
        self._oTextModel.MultiLine = True
        self._oTextModel.VScroll = True
        self._console_rectangle = Rectangle(100, 150, 250, 100)
        self.title("Console")

    def build(self) -> UnoControl:
        """
        @return: the executor
        """
        self._oDialog.setModel(self._oDialogModel)
        self._oDialogModel.insertByName("text", self._oTextModel)
        _set_rectangle(self._oDialogModel, self._console_rectangle)
        text_rectangle = Rectangle(
            MARGIN, MARGIN, self._console_rectangle.w - MARGIN * 2,
                            self._console_rectangle.h - MARGIN * 2
        )
        _set_rectangle(self._oTextModel, text_rectangle)
        return self._oDialog

    def title(self, title: str) -> "ConsoleDialogBuilder":
        """
        Set the title
        @param title: the title

        @return: self for fluent style
        """
        self._oDialogModel.Title = title
        return self

    def console_rectangle(
            self, x: int, y: int, w: int, h: int) -> "ConsoleDialogBuilder":
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



class ConsoleExecutorBuilder:
    """
    A ConsoleExecutorBuilder

    Example:
    ```
    builder = ConsoleExecutorBuilder()
    executor = builder.title("See the messages").autoclose(True).build()
    ```
    """

    def __init__(self):
        self._dialog_builder = ConsoleDialogBuilder()
        self._autoclose = False
        self._progress_handler = lazy(ConsoleHandler)

    def build(self) -> ConsoleExecutor:
        """
        @return: the executor
        """
        oDialog = self._dialog_builder.build()
        if self._progress_handler is None:
            progress_handler = StandardConsoleHandler(oDialog.getControl("text"))
        else:
            progress_handler = self._progress_handler

        return ConsoleExecutor(oDialog, self._autoclose, progress_handler)

    def handler(self, handler: ProgressHandler) -> "ConsoleExecutorBuilder":
        """
        Sets a specific handler for this executor
        @param handler: the handler
        @return: self for fluent style
        """
        self._progress_handler = handler
        return self

    def title(self, title: str) -> "ConsoleExecutorBuilder":
        """See ConsoleDialogBuilder.title"""
        self._dialog_builder.title(title)
        return self

    def autoclose(self, b: bool) -> "ConsoleExecutorBuilder":
        """
        Set the autoclose value : if False, don't close
        @param b: the value of autoclose
        @return:
        """
        self._autoclose = b
        return self

    def console_rectangle(
            self, x: int, y: int, w: int, h: int) -> "ConsoleExecutorBuilder":
        """See ConsoleDialogBuilder.rectangle"""
        self._dialog_builder.console_rectangle(x, y, w, h)
        return self


def trace_event(logname: str, enter_exit: bool=True
                ) -> Callable[[Callable], Callable]:
    """
    A decorator for methods of LibreOffice `XEventListener`s
    (see https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1lang_1_1XEventListener.html).

    ```
    class ActionListener(unohelper.Base, XActionListener):
        ...
        @trace_event(__name__)
        def actionPerformed(self, e):
            ... do something
    ```

    If "do something" fails, you'll have a trace in the log file with the
    call stack. Set the `enter_exit` parameter to False when using the
    decorator on event methods that might be called frequently
    (eg. `XTextListener.textChanged`) to avoid logging overhead.

    @param logname: the name of the log
    @param enter_exit: if `True`, then log when the control flow enters and
                        exits the event method. Default is `False`
    """
    logger = logging.getLogger(logname)

    if enter_exit:
        def _trace_event(func: Callable) -> Callable:
            @functools.wraps(func)
            def decorated_func(*args, **kwargs):
                _self = args[0]
                name = _self.__class__.__name__
                logger.debug("Enter %s.%s", name, func.__name__)
                try:
                    func(*args, **kwargs)
                except:
                    logger.exception("Exception")
                finally:
                    logger.debug("Exit %s.%s", name, func.__name__)

            return decorated_func
    else:
        def _trace_event(func: Callable) -> Callable:
            @functools.wraps(func)
            def decorated_func(*args, **kwargs):
                try:
                    func(*args, **kwargs)
                except:
                    logger.exception("Exception")

            return decorated_func

    return _trace_event