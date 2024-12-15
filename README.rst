|Build Status| |Code Coverage|

Py4LO (Python For LibreOffice)
==============================

Copyright (C) J. Férard 2016-2024

Py4LO is a simple toolkit to help you write and include Python macros in LibreOffice Calc spreadsheets.
Under GPL v.3

Overview
--------

The LibreOffice Basic is limited and Python is a far more powerful language to write macros.
Py4LO helps you to pack your Python macros in a LibreOffice Calc document and offers a small but useful
library to access LibreOffice objects.

Py4LO requires Python >= 3.5.

Features
--------
* Test Python macros, embed them in an existing LibreOffice document and open this document in one command line;
* Generate a debug new document from an existing Python macro
* Interface with Xray/MRI
* Helpers to access cells, add filters, create new documents, get used rows or data arrays
* Helpers to convert XNameAccess to dicts and XIndexAccess to lists
* Access ODS files content without opening them
* ...

Installation
------------

Needs Python 3.

Just ``git clone`` the repo:

.. code-block:: bash

    git clone https://github.com/jferard/py4lo.git

Then install requirements (you may need to be in adminstrator mode):

.. code-block:: bash

    pip3 install -r requirements.txt

For Ubuntu:

.. code-block:: bash

    sudo apt-get install libreoffice-script-provider-python

Quick start
-----------
(See the script in examples/quickstart)

Create a new `qs` dir and a `src/main` subdir:

.. code-block:: bash

    mkdir -p qs/src/main
    cd qs

Step 1
~~~~~~

Create a simple Python script ``qs.py`` :

.. code-block:: python

    # -*- coding: utf-8 -*-
    # py4lo: entry
    # py4lo: embed lib py4lo_typing
    # py4lo: embed lib py4lo_helper
    # py4lo: embed lib py4lo_dialogs
    from py4lo_dialogs import message_box

    def test(*args):
        message_box("A message", "py4lo")

Step 2
~~~~~~

Generate a debug document:

.. code-block:: bash

    python3 <py4lo dir>/py4lo init

Where ``<py4lo dir>`` points to the cloned repo. It will create a
``new-project.ods`` document with the Python ``test`` function attached
to a button.

Step 3
~~~~~~

Rename ``new-project.ods`` to ``qs.ods`` and edit the document if you
want. Add a title, move the button, change the styles, etc.

Step 4
~~~~~~

Create the ``qs.toml``:

.. code-block:: toml

    [src]
    source_ods_file = "./qs.ods"

Step 5
~~~~~~

Edit the Python script ``qs.py``:

.. code-block:: python

    # -*- coding: utf-8 -*-
    # py4lo: entry
    # py4lo: embed lib py4lo_typing
    # py4lo: embed lib py4lo_helper
    # py4lo: embed lib py4lo_dialogs
    from py4lo_dialogs import message_box

    def test(*args):
        message_box("Another message", "py4lo")

Step 6
~~~~~~

Update and test the new script:

.. code-block:: bash

    python3 <py4lo dir>/py4lo run


The library
-----------
The library contains the following modules:

- `py4lo_typing` provides basic typing support for UNO objects.
- `py4lo_helper` manipulate LO objects (cells, rows, sheets, ...).
- `py4lo_commons` provides some helpful methods and classes (a simple bus, access to a config file, ...) for Python objects (strs, lists, ...).
- `py4lo_io` read and write documents.
- `py4lo_ods` is useful to manipulate ods documents in pure Python. Document content is parsed as XML, and never opened with LO.
- `py4lo_dialogs`: create some useful dialogs.
- `py4lo_sqlite3`: use SQLite on Windows systems.

The lib modules are subject to the "classpath" exception of the GPLv3 (see https://www.gnu.org/software/classpath/license.html).


How to
------

Import in script A an object from script B
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

In ``scriptB.py``:

.. code-block:: python

    class O():
        ...

In ``scriptA.py``:

.. code-block:: python

    import scriptB
    o = O()

Import a library
~~~~~~~~~~~~~~~~

Py4LO provides several functions to ease the manipulation of LibreOffice
data structures. See below.

If you want to use those functions, you have to create an "entry" script:
* this script contains all the functions that are exposed through buttons
* this script uses some directives to tell Py4LO to do some initialization.

Example. In ``main.py`` (this is the "entry" script):

.. code-block:: python

    # py4lo: entry
    # py4lo: embed lib py4lo_helper

*Warning* The special object ``XSCRIPTCONTEXT`` of type
`com.sun.star.script.provider.XScriptContext <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1script_1_1provider_1_1XScriptContext.html>`_
is passed to the scripts called from LibreOffice, but not to the
imported modules. It's up to the script to pass this object to the
modules that need it.

**CAVEAT** If you have the LibreOffice quickstarter, new imports may not be recognized. You might have to kill manually the `soffice` process.

Notes:

* ``# py4lo: entry`` is a directive. This directive informs py4lo that the module is called from LibreOffice. This fixes the path so that the scripts are accessible
* ``# py4lo: embed lib py4lo_helper`` copies the library py4lo_ods.py in the ODS destination file and declare it as a script

SQLite databases and `DataArray` s
---------------------------------
Py4LO provides a module to work with SQLite databases, because the standard
Python `sqlite3` module is missing on Windows systems.
The `py4lo_sqlite3` module is low-level and does not comply with the
`PEP249 <https://peps.python.org/pep-0249/>`_, but it provides some useful
functions.

Raw storage
~~~~~~~~~~~
You may use a SQLite database to store a DataArray. Values of a DataArray are
strings, floats or Nones. Since this is raw data, SQL capacitues are not
really useful. You can:

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

Objects storage
~~~~~~~~~~~~~~~
A little more is having, for each ``MyObject``, a ``MyObjectHelper`` with some
methods:

* ``MyObjectHelper.from_data_row(data_row: DATA_ROW) -> MyObject`` to
  create a new object, with typed fields, from a data row.
* The reverse method ``MyObjectHelper.to_data_row(obj: MyObject) -> DATA_ROW``
  to create a data row from an object.
* ``MyObjectHelper.bind(stmt: Sqlite3Statement, obj: MyObject)`` to
  bind the fields of the object to the columns.
* The reverse method
  ``MyObjectHelper.from_sqlite_row(db_row: List[Any]) -> MyObject`` to
  create a new object from a sqlite row.

Now you are comfortable with handling the data:

.. code-block:: python

    objs = [
        MyObjectHelper.from_data_row(data_row)
        for data_row in data_array
    ]
    # do something with objects
    with sqlite_open(self._path, "rw") as db:
        with db.transaction():
            with db.prepare("INSERT INTO table VALUES(?, ?, ?, ?)") as stmt:
                for obj in objs:
                    stmt.reset()
                    MyObjectHelper.bind(stmt, obj)
                    stmt.execute_update()

Then you can work with your SQLite database, do complex queries:

.. code-block:: python

    with sqlite_open(self._path, "r") as db:
        with db.prepare("SELECT ... FROM table WHERE ... GROUP BY ... SORTED BY ...") as stmt:
            objs = [
                from_sqlite_row(db_row)
                for db_row in stmt.execute_query()
            ]

Working with complex objects
~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sometimes, you have to build objects on top of data rows. It's common when you
have a denormalized DataArray. In this case you have to build first records,
that are simple typed representations of the data rows. And then build object from
these records.

The design is roughly the same, with four methods to handle the records:

* ``MyRecordHelper.from_data_row(data_row: DATA_ROW) -> MyRecord``.
* ``MyRecordHelper.to_data_row(record: MyRecord) -> DATA_ROW``.
* ``MyRecordHelper.bind(stmt: Sqlite3Statement, obj: MyRecord)``.
* ``MyRecordHelper.from_sqlite_row(db_row: List[Any]) -> MyRecord``.

On top of these methods, you have
``MyObjectHelper.from_records(recs: List[MyRecord]) -> MyObject`` that builds
an object from records:

.. code-block:: python

    recs = [
        MyRecordHelper.from_data_row(data_row)
        for data_row in data_array
    ]
    rec_by_name = {}
    for rec in recs:
        rec_by_name.setdefault(rec.name, []).append(rec)

    objs = [
        MyObjectHelper.from_records(recs)
        for name, recs in rec_by_name.items()
    ]

You can then use an efficient storage for ``MyObject`` in the SQLite database.
See for instance https://en.wikipedia.org/wiki/Boyce%E2%80%93Codd_normal_form.

Test
----

From the py4lo directory:

.. code-block:: bash

   python3 -m pytest --cov-report term-missing --ignore=examples --ignore=megalinter-reports --cov=py4lo --cov=lib && python3 -m pytest --cov-report term-missing --ignore=examples --ignore=test --ignore=megalinter-reports --ignore=py4lo/__main__.py --cov-append --doctest-modules --cov=lib


.. |Build Status| image:: https://github.com/jferard/py4lo/actions/workflows/workflow.yml/badge.svg
    :target: https://github.com/jferard/py4lo/actions/workflows/workflow.yml
.. |Code Coverage| image:: https://codecov.io/github/jferard/py4lo/branch/master/graph/badge.svg
    :target: https://codecov.io/github/jferard/py4lo
