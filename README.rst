|Build Status| |Code Coverage|

Py4LO (Python For LibreOffice)
==============================

Copyright (C) J. Férard 2016-2020

Py4LO is a simple toolkit to help you write and include Python macros in LibreOffice Calc spreadsheets.
Under GPL v.3

Overview
--------

The LibreOffice Basic is limited and Python is a far more powerful language to write macros.
Py4LO helps you to pack your Python macros in a LibreOffice Calc document and offers a small but useful
library to access LibreOffice objects.

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

    > git clone https://github.com/jferard/py4lo.git

Then install requirements (you may need to be in adminstrator mode):

.. code-block:: bash

    > pip install -r requirements.txt

For Ubuntu:

.. code-block:: bash

    > sudo apt-get install libreoffice-script-provider-python

Quick start
-----------

Create a new dir:

.. code-block:: bash

    > mkdir mydir

Step 1
~~~~~~

Create a simple Python script ``myscript.py`` :

.. code-block:: python

    # -*- coding: utf-8 -*-
    # py4lo: entry
    # py4lo: embed lib py4lo_helper
    import py4lo_helper
    _ = py4lo_helper.Py4LO_helper.create(XSCRIPTCONTEXT)

    def test(*args):
        from com.sun.star.awt.MessageBoxType import MESSAGEBOX
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
        _.message_box(_.parent_win, "A message", "py4lo", MESSAGEBOX, BUTTONS_OK)

Step 2
~~~~~~

Generate a debug document:

.. code-block:: python

    > python <py4lo dir>/py4lo init

Where ``<py4lo dir>`` points to the cloned repo. It will create a
``new-project.ods`` document with the Python ``test`` function attached
to a button.

Step 3
~~~~~~

Rename ``new-project.ods`` to ``mydoc.ods`` and edit the document if you
want.

Step 4
~~~~~~

Create the ``py4lo.toml``:

.. code-block:: toml

    [src]
    source_file = "./mydoc.ods"

Step 5
~~~~~~

Edit the Python script ``myscript.py``:

.. code-block:: python

    # -*- coding: utf-8 -*-
    # py4lo: entry
    # py4lo: embed lib py4lo_helper
    import py4lo_helper
    _ = py4lo_helper.Py4LO_helper.create(XSCRIPTCONTEXT)

    def test(*args):
        from com.sun.star.awt.MessageBoxType import MESSAGEBOX
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
        _.message_box(_.parent_win, "Another message", "py4lo", MESSAGEBOX, BUTTONS_OK)

Step 6
~~~~~~

Update and test the new script:

.. code-block:: bash

    > python <py4lo dir>/py4lo test

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

Import in script A a library
~~~~~~~~~~~~~~~~~~~~~~~~~~~~

In ``scriptA.py``:

.. code-block:: python

    # py4lo: entry
    # py4lo: embed lib py4lo_helper
    import py4lo_helper
    _ = py4lo_helper.Py4LO_helper.create(XSCRIPTCONTEXT)

*Warning* The special object ``XSCRIPTCONTEXT`` of type
`\`com.sun.star.script.provider.XScriptContext <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1script_1_1provider_1_1XScriptContext.html>`__
is passed to the scripts called from LibreOffice, but not to the
imported modules. It's up to the script to pass this object to the
modules that need it.

**CAVEAT** If you have the LibreOffice quickstarter, new imports may not be recognized. You might have to kill manually the `soffice` process.

Notes:

* `# py4lo: entry` is a directive. This directive informs py4lo that the module is called from LibreOffice. This fixes the path so that the scripts are accessible
* `# py4lo: embed lib py4lo_helper` copies the library py4lo_ods.py in the ODS destination file and declare it as a Script

The library
-----------
The library is still limited:

- `py4lo_ods` is useful to manipulate ods documents in pure Python. Document content is parsed as XML, and never opened with LO.
- `py4lo_helper` manipulate LO objects (cells, rows, sheets, ...)
- `py4lo_commons` provides some helpful methods and classes (a simple bus, access to a config file, ...) for Python objects (strs, lists, ...).

The lib modules are subject to the "classpath" exception of the GPLv3 (see https://www.gnu.org/software/classpath/license.html).

Test
----

From the py4lo directory:

.. code-block:: bash

   py.test --ignore=example --cov-report term-missing --cov=py4lo --cov=lib && py.test --ignore=example --ignore=test --ignore=py4lo/__main__.py --cov-report term-missing --cov-append --doctest-modules --cov=py4lo --cov=lib

.. |Build Status| image:: https://travis-ci.org/jferard/py4lo.svg?branch=master
   :target: https://travis-ci.org/jferard/py4lo
.. |Code Coverage| image:: https://img.shields.io/codecov/c/github/jferard/py4lo/master.svg
   :target: https://codecov.io/github/jferard/py4lo?branch=master
