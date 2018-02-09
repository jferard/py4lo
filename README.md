[![Build Status](https://travis-ci.org/jferard/py4lo.svg?branch=master)](https://travis-ci.org/jferard/py4lo)
[![Code Coverage](https://img.shields.io/codecov/c/github/jferard/py4lo/master.svg)](https://codecov.io/github/jferard/py4lo?branch=master)

# Py4LO (Python For LibreOffice)
(C) J. Férard 2016-2018

Py4LO is a simple toolkit to help you write Python scripts for LibreOffice.

Under GPL v.3

## Overview
Py4LO helps you to pack your Python scripts in a LibreOffice Calc document, with a debug option. It also provides a mechanism to import objects from another script, and a small library to ease the use of LibreOffice services.

NB. The library is still very limited.

## Installation
Needs Python 3.

Just ```git clone``` the repo:

    > git clone https://github.com/jferard/py4lo.git

Then install requirements (you may need to be in adminstrator mode):

    > pip install -r requirements.txt

For Ubuntu:

    > sudo apt-get install libreoffice-script-provider-python

## Quick start
Create a new dir:

    > mkdir mydir

### Step 1
Create a simple Python script ```myscript.py``` :

    # -*- coding: utf-8 -*-
    # py4lo: import lib py4lo_helper
    _ = py4lo_helper.Py4LO_helper()

    def test(*args):
    	from com.sun.star.awt.MessageBoxType import MESSAGEBOX
    	from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
    	_.message_box(_.parent_win, "A message", "py4lo", MESSAGEBOX, BUTTONS_OK)

### Step 2
Generate a debug document:

    > python <py4lo dir>/py4lo init

Where ```<py4lo dir>``` points to the cloned repo. It will create a ```new-project.ods``` document with the Python ```test``` function attached to a button.

### Step 3
Rename ```new-project.ods``` to ```mydoc.ods``` and edit the document if you want.

### Step 4
Create the ```py4lo.toml```:

    source_file = "./mydoc.ods"

### Step 5
Edit the Python script ```myscript.py```:

    # -*- coding: utf-8 -*-
    # py4lo: import lib py4lo_helper
    _ = py4lo_helper.Py4LO_helper()

    def test(*args):
    	from com.sun.star.awt.MessageBoxType import MESSAGEBOX
    	from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
    	_.message_box(_.parent_win, "Another message", "py4lo", MESSAGEBOX, BUTTONS_OK)

### Step 6
Update and test the new script:

    > python <py4lo dir>/py4lo test

## How to
### Import in script A an object from script B
In ```scriptB.py```:

    class O():
    	...

In ```scriptA.py```:

    # py4lo: import scriptB
    o = O()

### Import in script A a library
In ```scriptA.py```:

    # py4lo: import lib py4lo_helper as _

*Warning*
The special object `XSCRIPTCONTEXT` of type [`com.sun.star.script.provider.XScriptContext](  https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1script_1_1provider_1_1XScriptContext.html) is passed to the scripts called from LibreOffice, but not to the imported modules. It's up to the script to pass this object to the modules that need it.

## Test
From the py4lo directory:

    python -m pytest test
