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
import datetime as dt
import logging
import time
import unittest
from unittest import mock
from webbrowser import get

import py4lo_dialogs
from py4lo_dialogs import (MessageBoxType, ExecutableDialogResults)
from py4lo_dialogs import (
    get_toolkit, message_box, input_box,
    place_widget, get_text_size, file_dialog, Size, FileFilter,
    folder_dialog, ProgressExecutorBuilder, StandardProgressHandler,
    ConsoleExecutorBuilder, StandardConsoleHandler, trace_event,
    set_uno_control_date, get_uno_control_date, set_uno_control_bool,
    get_uno_control_bool, set_uno_control_text, get_uno_control_text,
    set_uno_control_text_from_list, get_uno_control_text_as_list
)


class ToolkitTestCase(unittest.TestCase):
    @mock.patch('py4lo_dialogs.create_uno_service')
    def test_new_toolkit(self, create_uno_service):
        # arrange
        py4lo_dialogs._oToolkit = None
        service = mock.Mock()
        create_uno_service.side_effect = [service]

        # act
        ret = get_toolkit()

        # assert
        self.assertEqual(service, ret)
        self.assertEqual([mock.call('com.sun.star.awt.Toolkit')],
                         create_uno_service.mock_calls)

    def test_existing_toolkit(self):
        # arrange
        py4lo_dialogs._oToolkit = 1

        # act
        ret = get_toolkit()

        # assert
        self.assertEqual(1, ret)


class Py4LODialogsTestCase(unittest.TestCase):
    def test_place_widget(self):
        # prepare
        oModel = mock.Mock()

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
        oModel = mock.Mock()
        oDialogModel = mock.Mock()
        oDialogModel.createInstance.side_effect = [oModel]
        oControl = mock.Mock(MinimumSize=mock.Mock(Width=11, Height=17))
        us.side_effect = [oControl]

        # play
        actual_size = get_text_size(oDialogModel, "test")

        # verify
        self.assertEqual(Size(5.5, 8.5), actual_size)
        self.assertEqual("test", oModel.Label)
        self.assertEqual([
            mock.call.createInstance(
                'com.sun.star.awt.UnoControlFixedTextModel')
        ], oDialogModel.mock_calls)
        self.assertEqual([
            mock.call.setModel(oModel)
        ], oControl.mock_calls)

    @mock.patch("py4lo_dialogs.get_toolkit")
    @mock.patch("py4lo_dialogs.get_provider")
    def test_message_box(self, get_provider, gt):
        self.maxDiff = None

        # prepare
        toolkit = mock.Mock()
        gt.side_effect = [toolkit]

        # play
        message_box("title", "text")

        # verify
        self.assertEqual(toolkit.mock_calls, [
            mock.call.createMessageBox(get_provider().parent_win,
                                       MessageBoxType.MESSAGEBOX, 1,
                                       'title', 'text'),
            mock.call.createMessageBox().execute()
        ])

    @mock.patch("py4lo_dialogs.get_toolkit")
    def test_message_box_parent_win(self, gt):
        self.maxDiff = None

        # prepare
        toolkit = mock.Mock()
        gt.side_effect = [toolkit]
        pw = mock.Mock()

        # play
        message_box("title", "text", parent_win=pw)

        # verify
        self.assertEqual(toolkit.mock_calls, [
            mock.call.createMessageBox(pw,
                                       MessageBoxType.MESSAGEBOX, 1,
                                       'title', 'text'),
            mock.call.createMessageBox().execute()
        ])

    @mock.patch("py4lo_dialogs.create_uno_service")
    @mock.patch("py4lo_dialogs.get_toolkit")
    def test_input_box(self, gt, cus):
        self.maxDiff = None

        # prepare
        toolkit = mock.Mock()
        gt.side_effect = [toolkit]
        pw = mock.Mock(PosSize=mock.Mock(Width=10, Height=20))

        dialog = mock.Mock()
        dialog_model = mock.Mock()
        cus.side_effect = [dialog_model, dialog]
        dialog.getControl.side_effect = [mock.Mock(Text="foo")]

        # play
        ret = input_box("title", "text", parent_win=pw)

        # verify
        self.assertEqual("foo", ret)

    @mock.patch("py4lo_dialogs.create_uno_service")
    def test_file_dialog_single(self, us):
        # prepare
        ddir = "baz"
        oPicker = mock.Mock(SelectedFiles=["foo"])
        oPicker.execute.side_effect = [ExecutableDialogResults.OK]
        us.side_effect = [oPicker]

        # play
        actual = file_dialog("foo", [FileFilter("bar", "*.bar")], ddir)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertFalse(oPicker.MultiSelectionMode)
        self.assertEqual([
            mock.call.appendFilter('bar', '*.bar'), mock.call.execute(),
        ], oPicker.mock_calls)
        self.assertEqual("foo", actual)

    @mock.patch("py4lo_dialogs.create_uno_service")
    def test_file_dialog_single_no_filter(self, us):
        # prepare
        ddir = "baz"
        oPicker = mock.Mock(SelectedFiles=["foo"])
        oPicker.execute.side_effect = [ExecutableDialogResults.OK]
        us.side_effect = [oPicker]

        # play
        actual = file_dialog("foo", display_dir=ddir)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertFalse(oPicker.MultiSelectionMode)
        self.assertEqual([
            mock.call.execute(),
        ], oPicker.mock_calls)
        self.assertEqual("foo", actual)

    @mock.patch("py4lo_dialogs.create_uno_service")
    def test_file_dialog_single_none(self, us):
        # prepare
        ddir = "baz"
        oPicker = mock.Mock()
        oPicker.execute.side_effect = [ExecutableDialogResults.CANCEL]
        us.side_effect = [oPicker]

        # play
        actual = file_dialog("foo", [FileFilter("bar", "*.bar")], ddir)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertFalse(oPicker.MultiSelectionMode)
        self.assertEqual([
            mock.call.appendFilter('bar', '*.bar'), mock.call.execute()
        ], oPicker.mock_calls)
        self.assertIsNone(actual)

    @mock.patch("py4lo_dialogs.create_uno_service")
    def test_file_dialog_multiple(self, us):
        # prepare
        ddir = "baz"
        oPicker = mock.Mock(SelectedFiles=["foo", "bar"])
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
            mock.call.appendFilter('bar', '*.bar'), mock.call.execute(),
        ], oPicker.mock_calls)
        self.assertEqual(["foo", "bar"], actual)

    @mock.patch("py4lo_dialogs.create_uno_service")
    def test_file_dialog_multiple_empty(self, us):
        # prepare
        ddir = "baz"
        oPicker = mock.Mock()
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
            mock.call.appendFilter('bar', '*.bar'), mock.call.execute()
        ], oPicker.mock_calls)
        self.assertEqual([], actual)

    @mock.patch("py4lo_dialogs.create_uno_service")
    def test_folder_dialog(self, us):
        # prepare
        ddir = "baz"
        oPicker = mock.Mock(Directory="d")
        oPicker.execute.side_effect = [ExecutableDialogResults.OK]
        us.side_effect = [oPicker]

        # play
        actual = folder_dialog("foo", ddir)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertEqual([
            mock.call.execute()
        ], oPicker.mock_calls)
        self.assertEqual("d", actual)

    @mock.patch("py4lo_dialogs.create_uno_service")
    def test_folder_dialog_none(self, us):
        # prepare
        ddir = "baz"
        oPicker = mock.Mock()
        oPicker.execute.side_effect = [ExecutableDialogResults.CANCEL]
        us.side_effect = [oPicker]

        # play
        actual = folder_dialog("foo", ddir)

        # verify
        self.assertEqual("foo", oPicker.Title)
        self.assertEqual("baz", oPicker.DisplayDirectory)
        self.assertEqual([
            mock.call.execute()
        ], oPicker.mock_calls)
        self.assertIsNone(actual)


class ProgressExecutorTestCase(unittest.TestCase):
    @mock.patch("py4lo_dialogs.create_uno_service")
    def test_simple(self, us):
        # arrange
        py4lo_dialogs._oToolkit = None
        oBarModel = mock.Mock()
        oTextModel = mock.Mock()
        oDialogModel = mock.Mock()
        oDialogModel.createInstance.side_effect = [oBarModel, oTextModel]

        oBar = mock.Mock(Value=0)
        oBar.Model = mock.Mock(ProgressValueMin=0, ProgressValueMax=100)
        oText = mock.Mock()
        oDialog = mock.Mock()
        oDialog.getControl.side_effect = [oBar, oText]
        oTK = mock.Mock()
        us.side_effect = [oDialogModel, oDialog, oTK]

        pe = ProgressExecutorBuilder().build()

        def func(h: StandardProgressHandler):
            h.progress(10)
            h.message("foo")

        # act
        pe.execute(func)
        time.sleep(1)

        # assert
        self.assertEqual(us.mock_calls, [
            mock.call('com.sun.star.awt.UnoControlDialogModel'),
            mock.call('com.sun.star.awt.UnoControlDialog'),
            mock.call('com.sun.star.awt.Toolkit')
        ])
        self.assertEqual(oDialogModel.mock_calls, [
            mock.call.createInstance(
                'com.sun.star.awt.UnoControlProgressBarModel'),
            mock.call.createInstance(
                'com.sun.star.awt.UnoControlFixedTextModel'),
            mock.call.insertByName('bar', oBarModel),
            mock.call.insertByName('text', oTextModel),
        ])
        self.assertEqual(oDialog.mock_calls, [
            mock.call.setModel(oDialogModel),
            mock.call.getControl('bar'),
            mock.call.getControl('text'),
            mock.call.setVisible(True),
            mock.call.createPeer(oTK, None),
            mock.call.dispose()
        ])
        self.assertEqual(10, oBar.Value)
        self.assertEqual("foo", oText.Text)

    @mock.patch("py4lo_dialogs.create_uno_service")
    def test_build(self, us):
        # prepare
        py4lo_dialogs._oToolkit = None
        oBarModel = mock.Mock()
        oTextModel = mock.Mock()
        oDialogModel = mock.Mock()
        oDialogModel.createInstance.side_effect = [oBarModel, oTextModel]

        oBar = mock.Mock()
        oBar.Model = mock.Mock(ProgressValueMin=0, ProgressValueMax=100)
        oText = mock.Mock()
        oDialog = mock.Mock()
        oDialog.getControl.side_effect = [oBar, oText]
        oTK = mock.Mock()
        us.side_effect = [oDialogModel, oDialog, oTK]

        # play
        pe = ProgressExecutorBuilder().title(
            "bar").bar_dimensions(100, 50).autoclose(
            False).dialog_rectangle(10, 10, 120, 60).bar_progress(
            100, 1000).message("base msg").build()

        def func(h: StandardProgressHandler):
            h.progress(1000)
            h.message("foo")

        pe.execute(func)
        time.sleep(1)

        # verify
        self.assertEqual(us.mock_calls, [
            mock.call('com.sun.star.awt.UnoControlDialogModel'),
            mock.call('com.sun.star.awt.UnoControlDialog'),
            mock.call('com.sun.star.awt.Toolkit')
        ])
        self.assertEqual(oDialogModel.mock_calls, [
            mock.call.createInstance(
                'com.sun.star.awt.UnoControlProgressBarModel'),
            mock.call.createInstance(
                'com.sun.star.awt.UnoControlFixedTextModel'),
            mock.call.insertByName('bar', oBarModel),
            mock.call.insertByName('text', oTextModel),
        ])
        self.assertEqual(oDialog.mock_calls, [
            mock.call.setModel(oDialogModel),
            mock.call.getControl('bar'),
            mock.call.getControl('text'),
            mock.call.setVisible(True),
            mock.call.createPeer(oTK, None),
            mock.call.execute()
        ])
        self.assertEqual(100, oBar.Value)
        self.assertEqual("foo", oText.Text)


class ConsoleExecutorTestCase(unittest.TestCase):
    @mock.patch("py4lo_dialogs.create_uno_service")
    def test_simple(self, us):
        # arrange
        py4lo_dialogs._oToolkit = None
        oTextModel = mock.Mock()
        oDialogModel = mock.Mock()
        oDialogModel.createInstance.side_effect = [oTextModel]

        oText = mock.Mock()
        oDialog = mock.Mock()
        oDialog.getControl.side_effect = [oText]
        oTK = mock.Mock()
        us.side_effect = [oDialogModel, oDialog, oTK]

        ce = ConsoleExecutorBuilder().build()

        def func(h: StandardConsoleHandler):
            h.message("foo")

        # act
        ce.execute(func)
        time.sleep(1)

        # assert
        self.assertEqual([
            mock.call('com.sun.star.awt.UnoControlDialogModel'),
            mock.call('com.sun.star.awt.UnoControlDialog'),
            mock.call('com.sun.star.awt.Toolkit')
        ], us.mock_calls)
        self.assertEqual([
            mock.call.createInstance('com.sun.star.awt.UnoControlEditModel'),
            mock.call.insertByName('text', oTextModel),
        ], oDialogModel.mock_calls)
        self.assertEqual([
            mock.call.setModel(oDialogModel),
            mock.call.getControl('text'),
            mock.call.setVisible(True),
            mock.call.createPeer(oTK, None),
            mock.call.execute()
        ], oDialog.mock_calls)
        self.assertEqual(5, oTextModel.PositionX)
        self.assertEqual(5, oTextModel.PositionY)
        self.assertEqual(90, oTextModel.Height)
        self.assertEqual(240, oTextModel.Width)
        self.assertTrue(oTextModel.ReadOnly)
        self.assertTrue(oTextModel.MultiLine)
        self.assertTrue(oTextModel.VScroll)


class TraceEventTestCase(unittest.TestCase):
    def test_enter_exit(self):
        with self.assertLogs("foo", level=logging.DEBUG) as log:
            self.traced_one()

        self.assertEqual([
            'DEBUG:foo:Enter TraceEventTestCase.traced_one',
            'DEBUG:foo:Exit TraceEventTestCase.traced_one'
        ], log.output)

    def test_enter_raise_exit(self):
        with self.assertLogs("foo", level=logging.DEBUG) as log:
            self.traced_two()

        exc_lines = log.output[1].split("\n")

        self.assertEqual(3, len(log.output))
        self.assertEqual('DEBUG:foo:Enter TraceEventTestCase.traced_two',
                         log.output[0])
        self.assertEqual('ERROR:foo:Exception', exc_lines[0])
        self.assertEqual('Exception: bar', exc_lines[-1])
        self.assertEqual('DEBUG:foo:Exit TraceEventTestCase.traced_two',
                         log.output[2])

    def test_raise(self):
        with self.assertLogs("foo", level=logging.DEBUG) as log:
            self.traced_three()

        exc_lines = log.output[0].split("\n")

        self.assertEqual(1, len(log.output))
        self.assertEqual('ERROR:foo:Exception', exc_lines[0])
        self.assertEqual('Exception: bar', exc_lines[-1])

    @trace_event("foo")
    def traced_one(self):
        pass

    @trace_event("foo")
    def traced_two(self):
        raise Exception("bar")

    @trace_event("foo", enter_exit=False)
    def traced_three(self):
        raise Exception("bar")


class GetSetUnoDialogTestCase(unittest.TestCase):
    def test_set_uno_control_date_empty(self):
        # arrange
        oControl = mock.Mock()

        # act
        set_uno_control_date(oControl, None)

        # assert
        self.assertEqual([mock.call.setEmpty()], oControl.mock_calls)

    def test_set_uno_control_date_date(self):
        # arrange
        oControl = mock.Mock()

        # act
        set_uno_control_date(oControl, dt.date(2017, 11, 21))

        # assert
        self.assertEqual(2017, oControl.Date.Year)
        self.assertEqual(11, oControl.Date.Month)
        self.assertEqual(21, oControl.Date.Day)

    def test_get_uno_control_date_empty(self):
        # arrange
        oControl = mock.Mock()
        oControl.isEmpty.side_effect = [True]

        # act
        d = get_uno_control_date(oControl)

        # assert
        self.assertIsNone(d)

    def test_get_uno_control_date(self):
        # arrange
        struct = mock.Mock(Year=2003, Month=6, Day=3)
        oControl = mock.Mock(Date=struct)
        oControl.isEmpty.side_effect = [False]

        # act
        d = get_uno_control_date(oControl)

        # assert
        self.assertEqual(d, dt.date(2003, 6, 3))

    def test_set_uno_control_bool(self):
        # arrange
        oControl = mock.Mock()

        # act
        set_uno_control_bool(oControl, True)

        # assert
        self.assertEqual(1, oControl.State)

    def test_get_uno_control_bool(self):
        # arrange
        oControl = mock.Mock(State=True)

        # act
        b = get_uno_control_bool(oControl)

        # assert
        self.assertTrue(b)

    def test_set_uno_control_text_none(self):
        # arrange
        oControl = mock.Mock()

        # act
        set_uno_control_text(oControl, None)

        # assert
        self.assertEqual("", oControl.Text)

    def test_set_uno_control_text_blank(self):
        # arrange
        oControl = mock.Mock()

        # act
        set_uno_control_text(oControl, "  ")

        # assert
        self.assertEqual("", oControl.Text)

    def test_set_uno_control_text_blank_none(self):
        # arrange
        oControl = mock.Mock()

        # act
        set_uno_control_text(oControl, "  ", apply=None)

        # assert
        self.assertEqual("  ", oControl.Text)

    def test_set_uno_control_text_foo(self):
        # arrange
        oControl = mock.Mock()

        # act
        set_uno_control_text(oControl, " foo ")

        # assert
        self.assertEqual("foo", oControl.Text)

    def test_set_uno_control_text_foo2(self):
        # arrange
        oControl = mock.Mock()

        # act
        set_uno_control_text(oControl, " foo ", apply=str.upper)

        # assert
        self.assertEqual(" FOO ", oControl.Text)

    def test_set_uno_control_text_foo3(self):
        # arrange
        oControl = mock.Mock()

        # act
        set_uno_control_text(oControl, " foo ",
                             apply=lambda s: s.strip().capitalize())

        # assert
        self.assertEqual("Foo", oControl.Text)

    def test_get_uno_control_text_blank(self):
        # arrange
        oControl = mock.Mock(Text="  ")

        # act
        text = get_uno_control_text(oControl)

        # assert
        self.assertIsNone(text)

    def test_get_uno_control_text_blank_none(self):
        # arrange
        oControl = mock.Mock(Text="  ")

        # act
        text = get_uno_control_text(oControl, apply=None)

        # assert
        self.assertEqual("  ", text)

    def test_get_uno_control_text_void(self):
        # arrange
        oControl = mock.Mock(Text="")

        # act
        text = get_uno_control_text(oControl,
                                    apply=lambda x: x.strip().capitalize())

        # assert
        self.assertIsNone(text)

    def test_get_uno_control_text_foo(self):
        # arrange
        oControl = mock.Mock(Text=" foo ")

        # act
        text = get_uno_control_text(oControl)

        # assert
        self.assertEqual("foo", text)

    def test_get_uno_control_text_foo2(self):
        # arrange
        oControl = mock.Mock(Text=" foo ")

        # act
        text = get_uno_control_text(oControl, apply=None)

        # assert
        self.assertEqual(" foo ", text)

    def test_get_uno_control_text_foo3(self):
        # arrange
        oControl = mock.Mock(Text=" foo ")

        # act
        text = get_uno_control_text(oControl,
                                    apply=lambda x: x.strip().capitalize())

        # assert
        self.assertEqual("Foo", text)

    def test_set_uno_control_text_from_list(self):
        # arrange
        oControl = mock.Mock()

        # act
        set_uno_control_text_from_list(oControl, [" foo ", " ", "bar"])

        # assert
        self.assertEqual("foo\nbar", oControl.Text)

    def test_set_uno_control_text_from_list_none(self):
        # arrange
        oControl = mock.Mock()

        # act
        set_uno_control_text_from_list(
            oControl, [" foo ", " ", "bar"], delim=";", apply=None,
            filter_values=False
        )

        # assert
        self.assertEqual(' foo ; ;bar', oControl.Text)

    def test_get_uno_control_text_as_list(self):
        # arrange
        oControl = mock.Mock(Text=" foo\n  \nbar ")

        # act
        text = get_uno_control_text_as_list(oControl)

        # assert
        self.assertEqual(["foo", "bar"], text)

    def test_get_uno_control_text_as_list_none(self):
        # arrange
        oControl = mock.Mock(Text=" foo;  ;bar ")

        # act
        text = get_uno_control_text_as_list(oControl, delim=";", apply=None,
                                            filter_values=False)

        # assert
        self.assertEqual([' foo', '  ', 'bar '], text)


if __name__ == '__main__':
    unittest.main()
