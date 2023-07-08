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
import time
import unittest
from unittest import mock
from unittest.mock import Mock, call

from py4lo_dialogs import (MessageBoxType, ExecutableDialogResults)
from py4lo_dialogs import (message_box, place_widget,
                           get_text_size, file_dialog, Size, FileFilter,
                           folder_dialog, ProgressExecutorBuilder,
                           ProgressHandler, ConsoleExecutorBuilder,
                           ConsoleHandler)


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

    @mock.patch("py4lo_dialogs.create_uno_service")
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

    @mock.patch("py4lo_dialogs.create_uno_service_ctxt")
    @mock.patch("py4lo_dialogs.provider")
    def test_message_box(self, provider, usc):
        self.maxDiff = None

        # prepare
        toolkit = mock.Mock()
        usc.return_value = toolkit

        # play
        message_box("title", "text")

        # verify
        self.assertEqual([
            mock.call('com.sun.star.awt.Toolkit'),
            mock.call().createMessageBox(provider.parent_win,
                                         MessageBoxType.MESSAGEBOX, 1,
                                         'title', 'text'),
            mock.call().createMessageBox().execute()
        ], usc.mock_calls)

    @mock.patch("py4lo_dialogs.create_uno_service_ctxt")
    @mock.patch("py4lo_dialogs.provider")
    def test_message_box_parent_win(self, provider, usc):
        self.maxDiff = None

        # prepare
        toolkit = mock.Mock()
        usc.return_value = toolkit
        pw = Mock()

        # play
        message_box("title", "text", parent_win=pw)

        # verify
        self.assertEqual([
            mock.call('com.sun.star.awt.Toolkit'),
            mock.call().createMessageBox(pw,
                                         MessageBoxType.MESSAGEBOX, 1,
                                         'title', 'text'),
            mock.call().createMessageBox().execute()
        ], usc.mock_calls)

    @mock.patch("py4lo_dialogs.create_uno_service")
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

    @mock.patch("py4lo_dialogs.create_uno_service")
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

    @mock.patch("py4lo_dialogs.create_uno_service")
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

    @mock.patch("py4lo_dialogs.create_uno_service")
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

    @mock.patch("py4lo_dialogs.create_uno_service")
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

    @mock.patch("py4lo_dialogs.create_uno_service")
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

    @mock.patch("py4lo_dialogs.create_uno_service")
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
    @mock.patch("py4lo_dialogs.create_uno_service")
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
        oTK = Mock()
        us.side_effect = [oDialogModel, oDialog, oTK]

        # play
        pe = ProgressExecutorBuilder().build()

        def func(h: ProgressHandler):
            h.progress(10)
            h.message("foo")

        pe.execute(func)
        time.sleep(1)

        # verify
        self.assertEqual([
            call('com.sun.star.awt.UnoControlDialogModel'),
            call('com.sun.star.awt.UnoControlDialog'),
            call('com.sun.star.awt.Toolkit')
        ], us.mock_calls)
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
            call.createPeer(oTK, None),
            call.dispose()
        ], oDialog.mock_calls)
        self.assertEqual(10, oBar.Value)
        self.assertEqual("foo", oText.Text)

    @mock.patch("py4lo_dialogs.create_uno_service")
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
        oTK = Mock()
        us.side_effect = [oDialogModel, oDialog, oTK]

        # play
        pe = ProgressExecutorBuilder().title(
            "bar").bar_dimensions(100, 50).autoclose(
            False).dialog_rectangle(10, 10, 120, 60).bar_progress(
            100, 1000).message("base msg").build()

        def func(h: ProgressHandler):
            h.progress(1000)
            h.message("foo")

        pe.execute(func)
        time.sleep(1)

        # verify
        self.assertEqual([
            call('com.sun.star.awt.UnoControlDialogModel'),
            call('com.sun.star.awt.UnoControlDialog'),
            call('com.sun.star.awt.Toolkit')
        ], us.mock_calls)
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
            call.createPeer(oTK, None),
            call.execute()
        ], oDialog.mock_calls)
        self.assertEqual(1000, oBar.Value)
        self.assertEqual("foo", oText.Text)


class ConsoleExecutorTestCase(unittest.TestCase):
    @mock.patch("py4lo_dialogs.create_uno_service")
    def test_simple(self, us):
        # prepare
        oTextModel = Mock()
        oDialogModel = Mock()
        oDialogModel.createInstance.side_effect = [oTextModel]

        oText = Mock()
        oDialog = Mock()
        oDialog.getControl.side_effect = [oText]
        oTK = Mock()
        us.side_effect = [oDialogModel, oDialog, oTK]

        # play
        ce = ConsoleExecutorBuilder().build()

        def func(h: ConsoleHandler):
            h.message("foo")

        ce.execute(func)
        time.sleep(1)

        # verify
        # TODO: might not
        # self.assertEqual([
        #     call.insertText(ANY,  'foo\n')
        # ], oText.mock_calls)
        self.assertEqual([
            call('com.sun.star.awt.UnoControlDialogModel'),
            call('com.sun.star.awt.UnoControlDialog'),
            call('com.sun.star.awt.Toolkit')
        ], us.mock_calls)
        self.assertEqual([
            call.createInstance('com.sun.star.awt.UnoControlEditModel'),
            call.insertByName('text', oTextModel),
        ], oDialogModel.mock_calls)
        self.assertEqual([
            call.setModel(oDialogModel),
            call.getControl('text'),
            call.setVisible(True),
            call.createPeer(oTK, None),
            call.execute()
        ], oDialog.mock_calls)
        self.assertEqual(5, oTextModel.PositionX)
        self.assertEqual(5, oTextModel.PositionY)
        self.assertEqual(90, oTextModel.Height)
        self.assertEqual(240, oTextModel.Width)
        self.assertTrue(oTextModel.ReadOnly)
        self.assertTrue(oTextModel.MultiLine)
        self.assertTrue(oTextModel.VScroll)


if __name__ == '__main__':
    unittest.main()
