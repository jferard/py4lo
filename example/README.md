# Py4LO (Python For LibreOffice) Example
Py4LO is a simple toolkit to help you write Python scripts for LibreOffice.

## How to run the example?
Type in your command line interface:
```
> python ../py4lo run
```

It will update the ods file, run the tests and launch LibreOffice.

## How to see the script?
Type in your command line interface:
```
> python ../py4lo debug
```

Open the py4lo-debug.ods file.

It will update the ods file, run the tests and launch LibreOffice.

## How to mock objects if I have a prehistoric version of LibreOffice?
Just use flexmock. Type in your command line interface:
```
> python -m pip install flexmock
```

Why flexmock? Because the unittest.mock exists only since Python 3.3, that is LibreOffice 4.0. When you are stuck to an older version of LO, you can't mock anything without importing a external library. Flexmock is a good choice.
