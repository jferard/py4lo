# Py4LO (Python For LibreOffice)
Py4LO is a simple toolkit to help you write Python scripts for LibreOffice.

## Overview
Py4LO helps you to pack your Python scripts in a LibreOffice Calc document, with a debug option. It also provides a mechanism to import objects from another script, and a small library to ease the use of LibreOffice services.

## Installation
Needs Python 3.

Just ```git clone``` the repo:

```
> git clone https://github.com/jferard/py4lo.git
```

Then install requirements (you may need to be in adminstrator mode):
```
> pip install -r requirements.txt
```


## Quick start
Create a new dir:

```
> mkdir mydir
```

### Step 1
Create a simple Python script ```myscript.py``` :
```python
# -*- coding: utf-8 -*-
# py4lo: use lib "py4lo_helper::Py4LO_helper" as _
   
def test(*args):
	from com.sun.star.awt.MessageBoxType import MESSAGEBOX
	from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
	_.message_box(_.parent_win, "A message", "py4lo", MESSAGEBOX, BUTTONS_OK)
```

### Step 2
Generate a debug document:
```
> python <py4lo dir>/py4lo init
```

Where ```<py4lo dir>``` points to the cloned document. It will create a ```new-project.ods``` document with the Python ```test``` function attached to a button. 

### Step 3
Rename ```new-project.ods``` to ```mydoc.ods``` and edit the document if you want.

### Step 4
Create the ```py4lo.toml```:
```toml
source_file = "./mydoc.ods"
```

### Step 5
Edit the Python script ```myscript.py```:
```python
# -*- coding: utf-8 -*-
# py4lo: use lib "py4lo_helper::Py4LO_helper" as _
   
def test(*args):
	from com.sun.star.awt.MessageBoxType import MESSAGEBOX
	from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
	_.message_box(_.parent_win, "Another message", "py4lo", MESSAGEBOX, BUTTONS_OK)
```

### Step 6
Update and test the new script:
```
> python <py4lo dir>/py4lo test
```

## How to
### Import in script A an object from script B
In ```scriptB.py```:
```python
class O():
	...

def __export_o():
	return O()
```

In ```scriptA.py```:
```python
# py4lo: use "scriptB::o"
```

### Import in script A an object from a library
In ```scriptA.py```:
```python
# py4lo: use lib "py4lo_helper::Py4LO_helper" as _
```

NB. The library is still very limited.