|Build Status| |Code Coverage|

Py4LO (Python For LibreOffice)
==============================

Copyright (C) J. Férard 2016-2025

Py4LO is a simple toolkit to help you write and include Python macros in
LibreOffice Calc spreadsheets.
Under **GPL v.3**

Py4LO requires Python >= 3.8.

Overview
--------

The LibreOffice Basic is limited and Python is a far more powerful language
to write macros. Py4LO helps you to pack your Python macros in a LibreOffice
Calc or Writer document and offers a small but useful library to access
LibreOffice objects.

See also:

* https://github.com/jferard/py4lo/blob/master/QuickStart.rst
* https://github.com/jferard/py4lo/blob/master/From_LO_Basic_to_Python.rst

Features
--------
* Test Python macros, embed them in an existing LibreOffice document and open
  this document in one command line;
* Generate a debug new document from an existing Python macro
* Interface with Xray/MRI
* Helpers to access cells, add filters, create new documents, get used rows or data arrays
* Helpers to convert XNameAccess to dicts and XIndexAccess to lists
* Access ODS files content without opening them
* ...

Why Py4LO?
----------
When I created Py4LO, back in 2016, I did not know `APSO
<https://extensions.libreoffice.org/en/extensions/show/apso-alternative-script-organizer-for-python>`_,
the "Alternative Script Organizer for Python" (APSO was created in 2012, but
released on *extensions.libreoffice.org* in november 2016). Thus I had to
develop my own solution. I was programming in Java, and immediately tought of
something like `Maven <https://en.wikipedia.org/wiki/Apache_Maven>`_: a command
line tool that would test and compile my document with the macros. Py4LO was born.

Soon, I created a library to help me create new macros fasters. This design
might seem complex ("You mean I have to rerun the command each time
I modify a script?" - "Yes!"), but helps to keep the code clean and testable.

The library
-----------
The library contains the following modules:

- |py4lo_typing|_: basic typing support for UNO objects, and helps the
  IDE to check types.
- |py4lo_commons|_: some helpful methods and classes (a simple bus,
  access to a config file, ...) for Python objects (strs, lists, ...).
- |py4lo_helper|_: manipulate LO objects (cells, rows, sheets, ...).
- |py4lo_dialogs|_: create some useful dialogs.
- |py4lo_io|_: read and write Calc documents.
- |py4lo_ods|_: read ods documents in pure Python (document
  content is parsed as XML, and never opened with LO).
- |py4lo_base|_: work with LibreOffice Base documents.
- |py4lo_sqlite3|_: use SQLite on Windows systems.

The lib modules are subject to the "classpath" exception of the GPLv3 (see
https://www.gnu.org/software/classpath/license.html).

.. |py4lo_typing| replace:: ``py4lo_typing``
.. _py4lo_typing: https://github.com/jferard/py4lo/blob/master/lib/py4lo_typing.py

.. |py4lo_commons| replace:: ``py4lo_commons``
.. _py4lo_commons: https://github.com/jferard/py4lo/blob/master/lib/py4lo_commons.py

.. |py4lo_helper| replace:: ``py4lo_helper``
.. _py4lo_helper: https://github.com/jferard/py4lo/blob/master/lib/py4lo_helper.py

.. |py4lo_dialogs| replace:: ``py4lo_dialogs``
.. _py4lo_dialogs: https://github.com/jferard/py4lo/blob/master/lib/py4lo_dialogs.py

.. |py4lo_io| replace:: ``py4lo_io``
.. _py4lo_io: https://github.com/jferard/py4lo/blob/master/lib/py4lo_io.py

.. |py4lo_ods| replace:: ``py4lo_ods``
.. _py4lo_ods: https://github.com/jferard/py4lo/blob/master/lib/py4lo_ods.py

.. |py4lo_base| replace:: ``py4lo_base``
.. _py4lo_base: https://github.com/jferard/py4lo/blob/master/lib/py4lo_base.py

.. |py4lo_sqlite3| replace:: ``py4lo_sqlite3``
.. _py4lo_sqlite3: https://github.com/jferard/py4lo/blob/master/lib/py4lo_sqlite3.py

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


Test
----

From the py4lo directory:

.. code-block:: bash

   python3 -m pytest --cov-report term-missing --ignore=examples --ignore=megalinter-reports --cov=py4lo --cov=lib && python3 -m pytest --cov-report term-missing --ignore=examples --ignore=test --ignore=megalinter-reports --ignore=py4lo/__main__.py --cov-append --doctest-modules --cov=lib

.. |Build Status| image:: https://github.com/jferard/py4lo/actions/workflows/workflow.yml/badge.svg
    :target: https://github.com/jferard/py4lo/actions/workflows/workflow.yml
.. |Code Coverage| image:: https://codecov.io/github/jferard/py4lo/branch/master/graph/badge.svg
    :target: https://codecov.io/github/jferard/py4lo

