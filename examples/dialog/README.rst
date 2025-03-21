Py4LO (Python For LibreOffice) Dialog Example
======================================

Py4LO is a simple toolkit to help you write Python macros for
LibreOffice.

This example is a simple demonstration of Py4LO dialog helpers.

How to run the example?
-----------------------

Type in your command line interface:

.. code-block:: bash

    python3 ../../py4lo run

It will update the ods file, run the tests and launch LibreOffice.

How to see the script?
----------------------

Type in your command line interface:

.. code-block:: bash

    python3 ../../py4lo debug

Open the py4lo-debug.ods file.

Or:

.. code-block:: bash

    python3 ../../py4lo debug && libreoffice py4lo-debug.ods

All functions defined in example.py are available. This provides a base
document: move buttons, add styles, etc. (You might also find that some
functions are interesting.)