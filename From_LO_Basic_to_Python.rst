From LibreOffice Basic to Python
================================

This document presents a few hints to create application for LibreOffice
with Python macros using Py4LO. Some hints are not related to Py4LO.

The project structure
---------------------
Py4LO has a project structure inspired by Maven (the Apache tool "that can
manage a project's build, reporting and documentation").

The directory structure is::

    src/
       main/
       test/

The `main` directory must contain an entry point, typically `main.py`, and may
contain subdirectories. The
`three tier architecture <https://en.wikipedia.org/wiki/Multitier_architecture#Three-tier_architecture>`_
is a good way to isolate the application code from LibreOffice. We'll see that
this is a requirement to produce a testable and maintainable code. Let's create
this structure::

    src/
       main/
           application/
           data/
           domain/
           main.py
           presentation/
               dialog/
       test/
           test_application/
           test_data/
           test_domain/
           test_presentation/
               test_dialog/

The `dialog` directory is a part of the `presentation` tier that contains with
dialog factories.

The `domain` contains objects from the domain.

Quick start
-----------
Call a macro from a button
~~~~~~~~~~~~~~~~~~~~~~~~~~
Todo.

Get info about the API
~~~~~~~~~~~~~~~~~~~~~~
Todo.

Design
------
Error handling
~~~~~~~~~~~~~~
Todo.

Testing
~~~~~~~
Todo.

Presentation tier
-----------------
Create a dialog
~~~~~~~~~~~~~~~
You can build a dialog from scratch (`py4lo_dialogs` provides some functions
taht will help you).

You can also use ``provider.get_dialog("Standard.mydialog")`` to get a dialog
built with the LibreOffice dialog editor.

In basic, you assign macros to the dialog elements. In Python, it's a common
practice to programmatically add a listener:

.. code-block:: python

    def create_dialog(dialog_name: str) -> UnoControl:
        oDialog = provider.get_dialog("Standard.mydialog")

        oOkListener = OkListener(...)
        oOkButton = oDialog.getControl("ok_button")
        oOkButton.addActionListener(oOkListener)

See next section for more details about ``OkListener``.

A button listener
~~~~~~~~~~~~~~~~~
Todo.

Application tier
---------------
Todo.

Data tier
---------
Read a CSV file
~~~~~~~~~~~~~~~
In LO Basic, you have to load the file (``loadComponentFromURL``) and process
the DataArray. In Python, you can use the
`csv <https://docs.python.org/3/library/csv.html>`_
module to parse the file.

Assuming that each row represents a ``MyObject`` instance (``MyObject`` is
part of the `domain`), you can create
a ``MyObjectHelper`` with some useful methods. The first method is
``MyObjectHelper.from_row(row: List[str]) -> MyObject``. This method creates a
new object, with **typed** fields, from a row.

**It is very important build a
consistent Python object a soon as possible**: parse the dates, detect enum values,
parse booleans, decide whether a void value is ``None`` or an empty string.
Once the data is typed and the object built, you can work with it (and test it)
out of LibreOffice.

The reverse method, ``MyObjectHelper.to_row(obj: MyObject) -> List[str]``
may be useful to store an object in a CSV file.

Dealing with complex objects
^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Sometimes, you have to build objects on top of several rows. It's common when
you have a denormalized CSV file. In this case you have to build first records,
that are simple typed representations of the data rows. And then build object from
these records.

The design is roughly the same:
``MyRecordHelper.from_row(row: List[str]) -> MyRecord`` will build the
record. A record is a simple `dataclass <https://docs.python.org/3/library/dataclasses.html>_`
but, fields are typed. On top of these records, you have
``MyObjectHelper.from_records(recs: List[MyRecord]) -> MyObject`` that builds
an object from records:

.. code-block:: python

    recs = [
        MyRecordHelper.from_row(row)
        for row in csv_reader
    ]
    rec_by_name = {}
    for rec in recs:
        rec_by_name.setdefault(rec.name, []).append(rec)

    objs = [
        MyObjectHelper.from_records(recs)
        for name, recs in rec_by_name.items()
    ]

Read a DataArray
~~~~~~~~~~~~~~~~
A DataArray is an array of arrays of values mapped to a SheeCellRange (see
https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XCellRangeData.html and
https://wiki.documentfoundation.org/Documentation/DevGuide/Spreadsheet_Documents#Data_Array).
In Python, values of a DataArray are:

* ``str`` instances for text cells
* ``float`` instances for doubles, dates, hours, integers, boolean, percents, currencies, fractions.
* ``None`` for ``#N/A`` values.

The general idea is the same as when you `Process a CSV file`: build a
consistent Python object as soon as possible. You might have to build records before
you build objects.

Read a LibreOffice Base or a SQLite database
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Py4LO provides `py4lo_sqlite3` a module to work with SQLite databases, because
the standard Python `sqlite3` module is missing on Windows systems.
The `py4lo_sqlite3` module is low-level and does not comply with the
`PEP249 <https://peps.python.org/pep-0249/>`_, but it provides some useful
functions.

Py4LO provides also `py4lo_base`, a module to work with LibreOffice Base
documents.

The general idea is the same as when you `Process a CSV file`:  build a
consistent Python object as soon as possible. You might have to build records before
you build objects.

Write data to a CSV file or a DataArray
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Once objects or records are built, you'll need some methods to write them
into a CSV file, a DataArray or a database:

* ``MyObjectHelper.to_row(obj: MyObject) -> List[str]``
* ``MyObjectHelper.to_data_row(obj: MyObject) -> DATA_ROW``

Write data to a SQLite database
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Once you have Python objects, you can store them in one or several tables.

One table
^^^^^^^^^
A simple Python object may be stored in a table. Create a
``MyObjectHelper.bind(stmt: Sqlite3Statement, obj: MyObject)`` to
  bind the fields of the object to the columns of the table.

.. code-block:: python

    with sqlite_open(self._path, "rw") as db:
        with db.transaction():
            with db.prepare("INSERT INTO table VALUES(?, ?, ?, ?)") as stmt:
                for obj in objs:
                    stmt.reset()
                    MyObjectHelper.bind(stmt, obj)
                    stmt.execute_update()

Several tables
^^^^^^^^^^^^^^
When object are more than simple records, a minimal normalization (see for
instance  https://en.wikipedia.org/wiki/Boyce%E2%80%93Codd_normal_form)
is recommended. Use an abstract ``SQLBond`` class, with a
``SQLBond.bind(stmt: Sqlite3Statement)`` method. Each ``SQLBond`` is able
to bind variables to a statement. The method
``MyObjectHelper.table1_bonds(obj) -> List[SQLBond]`` returns a list
of bonds for the object:

.. code-block:: python

    with sqlite_open(self._path, "rw") as db:
        with db.transaction():
            with db.prepare("INSERT INTO table1 VALUES(?, ?, ?, ?)") as stmt:
                for obj in objs:
                    for bond in MyObjectHelper.table1_bonds(obj):
                        stmt.reset()
                        bond.bind(stmtobj)
                        stmt.execute_update()

Create as many ``MyObjectHelper.table<n>_bonds(obj) -> List[SQLBond]`` as
necessary.

Now that the objects are correctly stored, you can use the full power of SQL
queries. To handle the result of those queries,

Transfering data
~~~~~~~~~~~~~~~~
A classical need is to load data from a CSV file or a DataArray to a SQLite
database, or from a SQLite database to a DataArray or a CSV file.

If you come from LibreOffice Basic, you might think that keeping the
storage formats as close as possible is the best solution. It is not.
If you load a CSV file into a database, don't store values as ``TEXT`` in
the database. If you load a DataArray, don't store values as
``TEXT`` or ``DOUBLE`` in the database.

Why you shouldn't store a DataArray raw values into a SQLite database
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
You may use a SQLite database to store a DataArray. Values of a DataArray are
strings, floats or Nones. Since this is raw data, SQL capacities are not
really available. You can:

* ``SELECT * FROM <table>`` and put it back into a DataArray ;
* ``SELECT SUM(<column>) FROM <table>`` to check the sum of the values of a
  column ;
* do other basic checks.

But you can't:

* use the `SQLite date and time functions
  <https://www.sqlite.org/lang_datefunc.html>`_ since dates in a DataArray are
  a number of days since ``oDoc.NullDate``.
* rely on such a raw data to do complex queries: you'll need a more accurate
  typing:

  * ``NULL`` is ``#N/A``, but what about an empty string? Should it be treated
    as a ``NULL`` value?
  * What about bools or integers? They are mixed with floats.
  * How to use ``typeof(...)``?

Unless you use the database as a temporary storage or to do some basic check on
millions of data rows, you have to do a little more.

Why you should use `domain` objects to do the transfer
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
The solution is to build `domain` objects then store them:

.. code-block:: python

    objs = [
        MyObjectHelper.from_row(row)
        for row in csv_reader
    ]
    with sqlite_open(self._path, "rw") as db:
        with db.transaction():
            with db.prepare("INSERT INTO table VALUES(?, ?, ?, ?)") as stmt:
                for obj in objs:
                    stmt.reset()
                    MyObjectHelper.bind(stmt, obj)
                    stmt.execute_update()

This ensures that the database is a correct representation of the objects,
not of the raw data. This may seem overkill, but it has a lot of advantages:

* It comes for free because it uses functions that you have already written ;
* It allows you to check the values (e.g. `sum` of a column) ;
* SQL queries are easy to use.

Summary::

    DOMAIN                 +---> Objects ---+
                          /         ^        \
    ---------------------/----------|---------\----------------
                        /           |          \
                       /-------> Records        \
    DATA              /                           v
                  CSV File - // RAW DATA // -> SQLite database

If you try to use the bottom path (``RAW DATA``), you may experience some hard
times.

Ressources
----------
Todo.
