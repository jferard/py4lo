# -*- coding: utf-8 -*-
# py4lo: entry
# py4lo: embed lib py4lo_typing
# py4lo: embed lib py4lo_commons
# py4lo: embed lib py4lo_helper
# py4lo: embed lib py4lo_dialogs

from py4lo_dialogs import place_widget, ControlModel, Control, EventListener
from py4lo_helper import create_uno_service

try:
    # noinspection PyUnresolvedReferences
    from com.sun.star.awt import XActionListener
except ImportError:
    from _mock_objects import (  # type: ignore[assignment]
        XActionListener,  # pyright: ignore[reportGeneralTypeIssues]
    )


class HelloListener(EventListener, XActionListener):
    def __init__(self, oDialog):
        self._oDialog = oDialog

    def actionPerformed(self, _e):
        self._oDialog.Visible = False


def open_dialog(*args):
    oDialogModel = create_uno_service(ControlModel.Dialog)
    place_widget(oDialogModel, 300, 200, 150, 60)
    oDialogModel.Title = "Dialog Example"

    oLabelModel = oDialogModel.createInstance(ControlModel.FixedText)
    place_widget(oLabelModel, 10, 10, 130, 20)
    oLabelModel.Name = "HelloLabel"
    oLabelModel.Label = "Hello, world!"
    oDialogModel.insertByName(oLabelModel.Name, oLabelModel)

    oButtonModel = oDialogModel.createInstance(ControlModel.Button)
    place_widget(oButtonModel, 10, 30, 130, 20)
    oButtonModel.Name = "HelloButton"
    oButtonModel.Label = "Goodbye, world!"
    oDialogModel.insertByName(oButtonModel.Name, oButtonModel)

    oDialog = create_uno_service(Control.Dialog)
    oDialog.setModel(oDialogModel)

    oButton = oDialog.getControl(oButtonModel.Name)
    oButton.addActionListener(HelloListener(oDialog))

    oDialog.createPeer(create_uno_service("com.sun.star.awt.Toolkit"), None)
    oDialog.Visible = True
