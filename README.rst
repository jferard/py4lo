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

Test
----

From the py4lo directory:

.. code-block:: bash

   python3 -m pytest --cov-report term-missing --ignore=examples --ignore=megalinter-reports --cov=py4lo --cov=lib && python3 -m pytest --cov-report term-missing --ignore=examples --ignore=test --ignore=megalinter-reports --ignore=py4lo/__main__.py --cov-append --doctest-modules --cov=lib


.. |Build Status| image:: https://github.com/jferard/py4lo/actions/workflows/workflow.yml/badge.svg
    :target: https://github.com/jferard/py4lo/actions/workflows/workflow.yml
.. |Code Coverage| image:: https://codecov.io/github/jferard/py4lo/branch/master/graph/badge.svg
    :target: https://codecov.io/github/jferard/py4lo
