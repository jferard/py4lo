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
import unittest
from unittest import mock
from unittest.mock import Mock, call

from py4lo_dialogs import (message_box, MessageBoxType, place_widget, Size,
                           get_text_size, FileFilter, file_dialog,
                           folder_dialog, ProgressExecutorBuilder,
                           ProgressHandler)


class ExecutableDialogResults:
    from com.sun.star.ui.dialogs.ExecutableDialogResults import (
        OK, CANCEL)


class Py4LODialogsTestCase(unittest.TestCase):
    def test_place_widget(self):
        # prepare
        oModel = Mock()

        # play
        place_widget(oModel, 1, 2, 3, 4)

        # verify
        self.assertEqual(1, oModel.PositionX)
        self.assertEqual(2, oModel.PositionY)
        self.assertEqual(3, oModel.Width)
        self.assertEqual(4, oModel.Height)

    @mock.patch("py4lo_dialogs.uno_service")
    def test_get_text_size(self, us):
        # prepare
        oModel = Mock()
        oDialogModel = Mock()
        oDialogModel.createInstance.side_effect = [oModel]
        oControl = Mock(MinimumSize=Mock(Width=11, Height=17))
        us.side_effect = [oControl]

        # play
        actual_size = get_text_size(oDialogModel, "test")

        # verify
        self.assertEqual(Size(5.5, 8.5), actual_size)
        self.assertEqual("test", oModel.Label)
        self.assertEqual([
            call.createInstance('com.sun.star.awt.UnoControlFixedTextModel')
        ], oDialogModel.mock_calls)
        self.assertEqual([
            call.setModel(oModel)
        ], oControl.mock_calls)

    @mock.patch("py4lo_dialogs.uno_service_ctxt")
    @mock.patch("py4lo_dialogs.provider")
    def test_message_box(self, provider, uno_service_ctxt):
        self.maxDiff = None

        # prepare
        toolkit = mock.Mock()
        uno_service_ctxt.return_value = toolkit

        # play
        message_box("text", "title")

        # verify
        self.assertEqual([
            mock.call('com.sun.star.awt.Toolkit'),
            mock.call().createMessageBox(provider.parent_win,
                                         MessageBoxType.MESSAGEBOX, 1,
                                         'title', 'text'),
            mock.call().createMessageBox().execute()
        ], uno_service_ctxt.mock_calls)

    @mock.patch("py4lo_dialogs.uno_service_ctxt")
    @mock.patch("py4lo_dialogs.provider")
    def test_message_box_parent_win(self, provider, uno_service_ctxt):
        self.maxDiff = None

        # prepare
        toolkit = mock.Mock()
        uno_service_ctxt.return_value = toolkit
        pw = Mock()

        # play
        message_box("text", "title", parent_win=pw)

        # verify
        self.assertEqual([
            mock.call('com.sun.star.awt.Toolkit'),
            mock.call().createMessageBox(pw,
                                         MessageBoxType.MESSAGEBOX, 1,
                                         'title', 'text'),
            mock.call().createMessageBox().execute()
        ], uno_service_ctxt.mock_calls)

    @mock.patch("py4lo_dialogs.uno_service")
    def test_file_dialog_single(self, us):
        # prepare
        ddir = "baz"
        oPicker = Mock(SelectedFiles=["foo"])
        oPicker.execute.side_effect = [ExecutableDialogResults.OK]
        us.side_effect = [oPicker]

        # play
        actual = file_dialog("foo", [FileFilter("bar", "*.bar")], ddir)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertFalse(oPicker.MultiSelectionMode)
        self.assertEqual([
            call.appendFilter('bar', '*.bar'), call.execute(),
        ], oPicker.mock_calls)
        self.assertEqual("foo", actual)

    @mock.patch("py4lo_dialogs.uno_service")
    def test_file_dialog_single_no_filter(self, us):
        # prepare
        ddir = "baz"
        oPicker = Mock(SelectedFiles=["foo"])
        oPicker.execute.side_effect = [ExecutableDialogResults.OK]
        us.side_effect = [oPicker]

        # play
        actual = file_dialog("foo", display_dir=ddir)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertFalse(oPicker.MultiSelectionMode)
        self.assertEqual([
            call.execute(),
        ], oPicker.mock_calls)
        self.assertEqual("foo", actual)

    @mock.patch("py4lo_dialogs.uno_service")
    def test_file_dialog_single_none(self, us):
        # prepare
        ddir = "baz"
        oPicker = Mock()
        oPicker.execute.side_effect = [ExecutableDialogResults.CANCEL]
        us.side_effect = [oPicker]

        # play
        actual = file_dialog("foo", [FileFilter("bar", "*.bar")], ddir)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertFalse(oPicker.MultiSelectionMode)
        self.assertEqual([
            call.appendFilter('bar', '*.bar'), call.execute()
        ], oPicker.mock_calls)
        self.assertIsNone(actual)

    @mock.patch("py4lo_dialogs.uno_service")
    def test_file_dialog_multiple(self, us):
        # prepare
        ddir = "baz"
        oPicker = Mock(SelectedFiles=["foo", "bar"])
        oPicker.execute.side_effect = [ExecutableDialogResults.OK]
        us.side_effect = [oPicker]

        # play
        actual = file_dialog("foo", [FileFilter("bar", "*.bar")], ddir,
                             single=False)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertTrue(oPicker.MultiSelectionMode)
        self.assertEqual([
            call.appendFilter('bar', '*.bar'), call.execute(),
        ], oPicker.mock_calls)
        self.assertEqual(["foo", "bar"], actual)

    @mock.patch("py4lo_dialogs.uno_service")
    def test_file_dialog_multiple_empty(self, us):
        # prepare
        ddir = "baz"
        oPicker = Mock()
        oPicker.execute.side_effect = [ExecutableDialogResults.CANCEL]
        us.side_effect = [oPicker]

        # play
        actual = file_dialog("foo", [FileFilter("bar", "*.bar")], ddir,
                             single=False)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertTrue(oPicker.MultiSelectionMode)
        self.assertEqual([
            call.appendFilter('bar', '*.bar'), call.execute()
        ], oPicker.mock_calls)
        self.assertEqual([], actual)

    @mock.patch("py4lo_dialogs.uno_service")
    def test_folder_dialog(self, us):
        # prepare
        ddir = "baz"
        oPicker = Mock(Directory="d")
        oPicker.execute.side_effect = [ExecutableDialogResults.OK]
        us.side_effect = [oPicker]

        # play
        actual = folder_dialog("foo", ddir)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertEqual([
            call.execute()
        ], oPicker.mock_calls)
        self.assertEqual("d", actual)

    @mock.patch("py4lo_dialogs.uno_service")
    def test_folder_dialog_none(self, us):
        # prepare
        ddir = "baz"
        oPicker = Mock()
        oPicker.execute.side_effect = [ExecutableDialogResults.CANCEL]
        us.side_effect = [oPicker]

        # play
        actual = folder_dialog("foo", ddir)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertEqual([
            call.execute()
        ], oPicker.mock_calls)
        self.assertIsNone(actual)


class ProgressExecutorTestCase(unittest.TestCase):
    @mock.patch("py4lo_dialogs.uno_service")
    def test_simple(self, us):
        # prepare
        oBarModel = Mock()
        oTextModel = Mock()
        oDialogModel = Mock()
        oDialogModel.createInstance.side_effect = [oBarModel, oTextModel]

        oBar = Mock()
        oText = Mock()
        oDialog = Mock()
        oDialog.getControl.side_effect = [oBar, oText]
        us.side_effect = [oDialogModel, oDialog]

        # play
        pe = ProgressExecutorBuilder().build()

        def func(h: ProgressHandler):
            h.progress(10)
            h.message("foo")

        pe.execute(func)

        #
        self.assertEqual([
            call.createInstance('com.sun.star.awt.UnoControlProgressBarModel'),
            call.createInstance('com.sun.star.awt.UnoControlFixedTextModel'),
            call.insertByName('bar', oBarModel),
            call.insertByName('text', oTextModel),
        ], oDialogModel.mock_calls)
        self.assertEqual([
            call.setModel(oDialogModel),
            call.getControl('bar'),
            call.getControl('text'),
            call.setVisible(True),
            call.dispose()
        ], oDialog.mock_calls)
        self.assertEqual(10, oBar.Value)
        self.assertEqual("foo", oText.Text)

    @mock.patch("py4lo_dialogs.uno_service")
    def test_build(self, us):
        # prepare
        oBarModel = Mock()
        oTextModel = Mock()
        oDialogModel = Mock()
        oDialogModel.createInstance.side_effect = [oBarModel, oTextModel]

        oBar = Mock()
        oText = Mock()
        oDialog = Mock()
        oDialog.getControl.side_effect = [oBar, oText]
        us.side_effect = [oDialogModel, oDialog]

        # play
        pe = ProgressExecutorBuilder().title(
            "bar").bar_dimensions(100, 50).autoclose(
            False).dialog_rectangle(10, 10, 120, 60).bar_progress(
            100, 1000).message("base msg").build()

        def func(h: ProgressHandler):
            h.progress(1000)
            h.message("foo")

        pe.execute(func)

        #
        self.assertEqual([
            call.createInstance('com.sun.star.awt.UnoControlProgressBarModel'),
            call.createInstance('com.sun.star.awt.UnoControlFixedTextModel'),
            call.insertByName('bar', oBarModel),
            call.insertByName('text', oTextModel),
        ], oDialogModel.mock_calls)
        self.assertEqual([
            call.setModel(oDialogModel),
            call.getControl('bar'),
            call.getControl('text'),
            call.setVisible(True),
            call.execute()
        ], oDialog.mock_calls)
        self.assertEqual(1000, oBar.Value)
        self.assertEqual("foo", oText.Text)


if __name__ == '__main__':
    unittest.main()
