# see http://stackoverflow.com/questions/61151/where-do-the-python-unit-tests-go
import sys
from pathlib import Path

test_dir = str(Path(__file__).parent.parent.parent)
if test_dir not in sys.path:
    sys.path.insert(0, test_dir)
from test_helper import *
