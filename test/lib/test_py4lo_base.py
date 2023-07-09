import unittest
from pathlib import Path
from unittest import mock
from py4lo_base import open_or_create_db, DataType


class Py4LOBaseTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self._oDB = mock.Mock()
        self._oDBContext = mock.Mock()
        self._oHandler = mock.Mock()
        self._oConnection = mock.Mock()
        self._oStatement = mock.Mock()

        self._oDB.connectWithCompletion.side_effect = [self._oConnection]
        self._oConnection.createStatement.side_effect = [self._oStatement]

    def tearDown(self) -> None:
        pass

    @mock.patch("py4lo_base.create_uno_service")
    def test_execute_update(self, cus):
        cus.side_effect = [self._oDBContext, self._oHandler]
        self._oDBContext.getByName.side_effect = [self._oDB]

        db = open_or_create_db(Path("./test.odb"))
        db.execute_update("SQL")

        self.assertEqual(
            [mock.call.executeUpdate('SQL')], self._oStatement.mock_calls)

    @mock.patch("py4lo_base.create_uno_service")
    def test_table_builder(self, cus):
        oTables = mock.Mock()
        oTableDataDescriptor = mock.Mock()
        oColumns = mock.Mock()
        oColDataDescriptor = mock.Mock()

        cus.side_effect = [self._oDBContext, self._oHandler]
        self._oDBContext.getByName.side_effect = [self._oDB]
        self._oConnection.getTables.return_value = oTables
        oTables.hasByName.return_value = False
        oTables.createDataDescriptor.return_value = oTableDataDescriptor
        oTableDataDescriptor.getColumns.return_value = oColumns
        oColumns.createDataDescriptor.return_value = oColDataDescriptor

        db = open_or_create_db(Path("./test.odb"))
        builder = db.get_table_builder("ok")
        builder.add_column("c", DataType.NUMERIC)
        builder.build()

        self.assertEqual([
            mock.call.hasByName('ok'),
            mock.call.createDataDescriptor(),
            mock.call.createDataDescriptor().getColumns(),
            mock.call.createDataDescriptor().getColumns().createDataDescriptor(),
            mock.call.createDataDescriptor().getColumns().appendByDescriptor(
                oColDataDescriptor),
            mock.call.appendByDescriptor(oTableDataDescriptor)
        ], oTables.mock_calls)


if __name__ == '__main__':
    unittest.main()
