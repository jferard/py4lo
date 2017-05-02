# see http://stackoverflow.com/questions/61151/where-do-the-python-unit-tests-go
import sys
import os

# append module root directory to sys.path
test_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
root_dir = os.path.dirname(test_dir)
py4lo_dir = os.path.join(root_dir, "py4lo")
lib_dir = os.path.join(root_dir, "lib")
inc_dir = os.path.join(root_dir, "inc")

sys.path.extend([py4lo_dir, lib_dir, inc_dir])
