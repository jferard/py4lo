# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016 J. Férard <https://github.com/jferard>
  
   This file is part of Py4LO.
  
   FastODS is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.
  
   FastODS is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.
  
   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""
import sys
import argparse
from cmd import test, debug, update, init

sys.argv.pop(0) # remove the 
command = sys.argv[0]

if command == '-h' or command == '--help':
	print ("""usage: py4lo.py -h|--help|command [args]

Python for LibreOffice.

-h, --help  show this help message and exit
command     a command = debug|help|init|test|update
            debug:  creates a debug.ods file with button for each function
			help:   more specific help
			init:   create a standard file
			test:   update + open the created file
			update: updates the file with all scripts""")
elif command == 'test':
	parser = argparse.ArgumentParser(description='Python for LibreOffice : test')
	args = parser.parse_args()
	test()
elif command == 'debug':
	parser = argparse.ArgumentParser(description='Python for LibreOffice : debug')
	args = parser.parse_args()
	debug()
elif command == 'init':
	parser = argparse.ArgumentParser(description='Python for LibreOffice : debug')
	args = parser.parse_args()
	init()
elif command == 'update':
	parser = argparse.ArgumentParser(description='Python for LibreOffice : update')
	args = parser.parse_args()
	update()
elif command == 'help':
	print("""TODO. See https://github.com/jferard/py4lo/README.md""")