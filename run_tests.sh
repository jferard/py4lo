#!/bin/bash

clear
ruff check
source .venv/bin/activate
export PYTHONPATH="`pwd`/lib/":"`pwd`/py4lo/":"/usr/lib/python3/dist-packages/"
echo "$PYTHONPATH"
python -c "import sys as s; print(s.path)"
python -m pytest --cov-report term-missing --ignore=examples --cov=py4lo --cov=lib --cov-report=xml
python -m pytest --cov-report term-missing --ignore=examples --ignore=test --ignore="py4lo/__main__.py" --cov-append --doctest-modules --cov=lib --cov-report=xml
deactivate
