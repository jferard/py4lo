Py4LO Quick start
=================
See the script in `examples/quickstart <https://github.com/jferard/py4lo/tree/master/examples/quickstart>`_

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

Import a module from Py4LO library
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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

**CAVEAT** If you have the LibreOffice quickstarter, new imports may not be
recognized. You might have to kill manually the `soffice` process.

Notes:

* ``# py4lo: entry`` is a directive. This directive informs py4lo that the module is called from LibreOffice. This fixes the path so that the scripts are accessible
* ``# py4lo: embed lib py4lo_helper`` copies the library py4lo_ods.py in the ODS destination file and declare it as a script

Import pure Python module
~~~~~~~~~~~~~~~~~~~~~~~~~
Put the module into the `src/opt` directory. Then use the ``embed script``
directive to embed it into the document:

.. code-block:: python

    # py4lo: embed script my_script



