From LibreOffice Basic to Python
================================

This document presents a few hints to create application for LibreOffice
with Python macros using Py4LO. Some hints are not related to Py4LO.

Working with Python
-------------------
The Python language
~~~~~~~~~~~~~~~~~~~
The Python language is way richer than Basic. But the main difference is
the Python standard library. On Windows platforms, this library is restricted
but still contains a lot of interesting modules. Those module provide
functionalities that are out of reach of Basic.

Note : On Windows platforms, the |sqlite3_module|_ is missing, but Py4LO
provides a workaround: the ``py4lo_sqlite3`` module.

Some of the differences:

* Python has a rich standard library whereas LibreOffice Basic has the
  ``Tools`` module and the ``ScriptForge`` library;
* Python is an object oriented programming language whereas Basic is not, even
  if you can cheat with ``Option Compatible``;
* Python has functional programming abilities (functions are variables), Basic
  has not;
* Python has a powerful exception mechanism, whereas Basic has ``On Error
  GoTo``;
* Python has standard data structures (lists, sets, dicts, heaps...) whereas
  Basic has arrays and poor ``Collections`` (or the EnumerableMap API from
  LibreOffice);
* Python provides plain text, XML and JSON input/output functions whereas
  Basic is limited;
* Python has a powerful logging module, whereas Basic has ``MessageBox``;
* Python is supported by IDEs (PyCharm is my favourite).

.. |sqlite3_module| replace:: ``sqlite3`` module
.. _sqlite3_module: https://docs.python.org/3/library/sqlite3.html


Python his very (case) sensitive
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
The API is the same for Python or Basic. If you know how to write macros
in Basic, you'll see that most of the code will be very similar. But when you
create macros using Python, you need to pay more attention to the API,
since the methods and fields are case sensitive.

Something like:

.. code-block:: python

    oSheet = oDoc.currentcontroller

wont' work, because nor the field ``currentcontroller`` neither the function
``getcurrentcontroller()`` exists in the API. But |getcurcontroller|_
(and the alias ``CurrentController``) exists:

.. code-block:: python

    oSheet = oDoc.CurrentController

Will work.

.. |getcurcontroller| replace:: ``getCurrentController()``
.. _getcurcontroller: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XModel.html#a44c3b26a1116ab41654d60357ccda9e1


Calling a function in Python
~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If the fnuction has no parameter, don't forget the parenthesis:

.. code-block:: python

    s = my_function

will assign the *function* to ``s``: ``s`` is now a function, not the result
of a function. Always add parenthesis:

.. code-block:: python

    s = my_function()

Managing resources in Python
~~~~~~~~~~~~~~~~~~~~~~~~~~~~
The typical form in standard Python is:

.. code-block:: python

    r = get_resources()
    try:
        do_wathever_you_want(r)
    except Exception as e:
        ...
    finally:
        free_resources(r)

When possible, you should use the ``with`` statement (see
`PEP 343 <https://peps.python.org/pep-0343/>`_ and |contextmanager|_
that ensures that the resources are freed).

.. code-block:: python

    with get_resources() as r:
        try:
            do_wathever_you_want(r)
        except Exception as e:
            ...


.. |contextmanager| replace:: ``@contextmanager``
.. _contextmanager: https://docs.python.org/3/library/contextlib.html#contextlib.contextmanager

EAFP (Easier to ask for forgiveness than permission)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Pytohn has a lot of exceptions that are part of the control flow. These are
not errors. Let's have a look at an example:

.. code-block:: python

    try:
        i = s.index("foo")
    except ValueError:
        ... # "foo" was not found
    else:
        ... # "foo" was found at index i

Is the same as:

.. code-block:: python

    i = s.find("foo")
    if i == -1
        ... # "foo" was not found
    else:
        ... # "foo" was found at index i

Python uses a lot of control flow exceptions, because of the
`EAFP <https://docs.python.org/3/glossary.html#term-EAFP>`_ philosophy. Common
examples of "EAFP"s: casts (``int``, ``float``...), search in strings or
sequences, use of attributes that may not exist, decode from bytes to string...

Logging
~~~~~~~
Python provides a simple |logging_module|_. You first need to configure the
logger, then use it like this:

.. code-block:: python

    class Foo:
        _logger = logging.getLogger(__name__)

        ...

        def f(self, x: int):
            self._logger.debug("f(%s) was called", i)
            ...

.. |logging_module| replace:: ``logging`` module
.. _logging_module: https://docs.python.org/3/library/logging.html

Python and LibreOffice
----------------------
The project structure
~~~~~~~~~~~~~~~~~~~~~
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

Call a macro from a button
~~~~~~~~~~~~~~~~~~~~~~~~~~
In Basic, you create a command button using the **Forms Controls**. See
`Adding a Command Button to a Document <https://help.libreoffice.org/latest/en-GB/text/shared/guide/formfields.html>`_.
But the LibreOffice interface does not provide a simple way to assign a Python
macro to the button.

Py4LO
^^^^^
Py4LO provides a way to create button from functons. If you run the
``py4lo init`` command, Py4Lo will read the sources files to find functions,
and create a button for each function. Take the ``new-project.ods`` document
and start with this document.

Manually
^^^^^^^^
Otherwise, you'll have to edit the XML file ``content.xml``.

First, create the button and assign whatever Basic macro you want to the
"execute" action. LibreOffice will set an URL for the action (see
`com.sun.star.uri.XVndSunStarScriptUrl <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1uri_1_1XVndSunStarScriptUrl.html>`_
for details on the URL format).

Open the ODS file with an utility to decompress zip files (Ark, File Roller,
7-Zip...). You'll the ``content.xml`` file at the root level. Edit this file
and search the string "vnd.sun.star.script".

You will find a sequence like this one:

.. code-block:: xml

    <script:event-listener script:language="ooo:script" script:event-name="form:performaction" xlink:href="vnd.sun.star.script:[...]?language=Basic&amp;location=document" xlink:type="simple"/>

You have to change the ``xlink:href`` attribute to something like:

.. code-block:: xml

    "vnd.sun.star.script:main.py$function?language=Python&amp;location=document"

and put the modified ``content.xml`` file into the archive.

Get info about the API
~~~~~~~~~~~~~~~~~~~~~~
The main information source about the API is the `LibreOffice SDK API
Reference <https://api.libreoffice.org/docs/idl/ref/index.html>`_. Bookmark
this link and use it!

There are two extensions that you might have used in Basic: `XRay
<https://wiki.openoffice.org/w/images/c/c6/XrayTool60_en.odt>`_ and `MRI
<https://extensions.libreoffice.org/en/extensions/show/mri-uno-object-inspection-tool>`_.

Py4LO
^^^^^
Py4LO provides the functions ``xray`` and ``mri`` (see ``py4lo_helper`` module).
Those functions will fail if the extensions are not installed.

Manually
^^^^^^^^
Here's a code snippet to use XRay:

.. code-block:: python

    oScriptProvider = oDoc.getScriptProvider()
    oScript = oScriptProvider.getScript(
                "vnd.sun.star.script:XrayTool._Main.Xray?"
                "language=Basic&location=application")
    oScript.invoke((obj,), tuple(), tuple())

Error handling
~~~~~~~~~~~~~~
Error handling and resources management is quite cumbersome in Basic
(``On Error GoTo ...`` is not easy to use).

Now that, coming from Basic, you discovered the ``try`` / ``except`` statement,
you imagine that you'll be able to control everything by fixing errors as soon
as they are raised. That may be optimistic.

We've alreay seen exceptions that are part of the control flow (see EAFP).
Those exceptions are handled as soon as they are raised. But what about other
exceptions (actual errors: network not available, not enough disk space, data
of unexpected type...)?

Usually, the most sensible thing to do is to log errors to be able to
understand what went wrong and quit gracefully. That may seem frustrating.

You are not convinced? You want your function that writes a 1kb file to catch
`OSError <https://docs.python.org/3/library/exceptions.html#OSError>`_,
`MemoryError <https://docs.python.org/3/library/exceptions.html#MemoryError>`_
and other kind of errors. Who knows? That might happen and you are never too
careful.

Okay. Think of responsibility. Is a tiny function responsible for the
disk space, the memory and a bunch of other "big" things that might go wrong?
You'd think that is the OS that is responsible for that, not the function!
And you can't, most of the time, fix the unexpected situation: it's too late.
Too late for asking the user to correct the data, free some disk space or
something like that.

**Rule of thumb**: If you think that a function is responsible for handling
an error, ask yourself if the calling function is not *more responsible*
than your function.

If you follow this rule of thumb, then you'll discover that the top function
is often responsible for handling the errors. Thus, always wrap your entry
points (macro assigned to a function or event listeners) with this code:

.. code-block:: python

    try:
        ...
    except Exception as e:
        self._logger.exception("Something bad happened!")
        message_box("Error", "Contact me and send the log file")

Testing
~~~~~~~
Python offers a powerful unittest_module_. You can mock objects, including
LibreOffice API objects, to test your code.

.. |unittest_module| replace:: ``unittest`` module
.. _unittest_module: https://docs.python.org/3/library/unittest.html

Debugging
~~~~~~~~~
In Basic, the switch between the IDE and the running macro is very easy: as
soon as an error is raised, you are in the IDE at the line of code, and you
can fix it.

In Python, you will have a murky message and no way to edit the code in place.
Therefore, you'll have to limited the back and forth between the code and the
execution.

The solution is to avoid debugging: test and log.

Py4LO configures the logger.

Presentation tier
-----------------
Create a dialog
~~~~~~~~~~~~~~~
You can build a dialog from scratch (``py4lo_dialogs`` provides some functions
that will help you).

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

Understanding threads
~~~~~~~~~~~~~~~~~~~~~
LibreOffice calls a Python script when:

* the user calls a macro from **Tools > Macros > Run macro** menu ;
* the user hits a button having a linked macro ;
* the user selects an item linked to a macro from a custom menu;
* the user triggers an event in a dialog.

When the Python script starts, LibreOffice is interrupted, waiting for
the end of the script.

We may have the following sequence::

                     user hits                                                   user hits the
                     a LO button                        the dialog is shown      dialog button
    LibreOffice:  ------|                             |-------------------------------|
                        | <-- linked script is called |                               | <--- listener calls Python script.
    Python     :        |-----------------------------|                               |------------------------
                          script creates a dialog,                                      python script
                          adds a button with a
                          listener

But what if I want to interact with LibreOffice component (document, sheets,
cells, infobar...) while the script is running? This won't have any visible
effect until the script ends and LibreOffice takes back the control.

Try this function:

.. code-block:: python

    def func(*_args):
        oDoc = provider.doc
        oSheet = oDoc.CurrentController.ActiveSheet
        oSheet.getCellByPosition(0, 0).String = "foo"
        time.sleep(1.0)
        oSheet.getCellByPosition(0, 0).String = "bar"
        time.sleep(1.0)
        oSheet.getCellByPosition(0, 0).String = "baz"
        time.sleep(1.0)

You may think that the "foo", "bar", baz" string will appear successively in
cell **A1** of the current active sheet instantly, but it won't.
It will take 3 seconds to leave the Python script and during this 3 seconds,
nothing will change in the sheet (
LibreOffice "hangs"). Then the last string, "baz", will appear in the cell **A1**.

Side note: if you understand this, you won't use
`oDoc.lockControllers() <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XModel.html#a7b7d36374033ee9210ec0ac5c1a90d9f>`_
and `oDoc.unlockControllers() <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XModel.html#abc62472c203de4d1403802509b153270>`_
anymore in Python: the interface is already locked.

There is a reason why you might want to update the LibreOffice components
during a script : when running a long script, you need to inform the user
of what's happening. Furthermore, we don't like when LibreOffice hangs for a
long time.

There is a solution: Python threads. If we start a thread in the Python script,
but do not wait until the thread finished (ie do not use `Thread.join() <https://docs.python.org/3/library/threading.html#threading.Thread.join>`_),
then LibreOffice will take the control back, but the Python thread will
continue tu be executed::

                     user hits                    main script      end of                                        end of
                     a LO button                creates a thread   main script                                thread script
    LibreOffice:  ------|                             |              |------------------ ... -----------------------|-------------------
                        | <-- linked script is called |              |                     ^                        |
    Python     :        |-----------------------------|--------------|            updates LibreOffice  component    |
                                  main script         |                                    |                        |
    Python thread :                                   |--------------------------------- ... -----------------------|
                                                                             thread script

This function will give the expected result (write "foo", wait 1 second, write
"bar", wait 1 second, write "baz", wait 1 second):

.. code-block:: python

    def func(*_args):
        oDoc = provider.doc
        oSheet = oDoc.CurrentController.ActiveSheet

        def aux():
            oSheet.getCellByPosition(0, 0).String = "foo"
            time.sleep(1.0)
            oSheet.getCellByPosition(0, 0).String = "bar"
            time.sleep(1.0)
            oSheet.getCellByPosition(0, 0).String = "baz"
            time.sleep(1.0)

        t = threading.Thread(target=aux)
        t.start()

Use this method if you want to update a dialog, an infobar, etc.

Beware: if there is an error in a Python thread, LibreOffice won't
show any error message.

Application tier
----------------
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
