#!/usr/bin/env python

import sys
import os
import subprocess

import re
import fnmatch
import csv
import textwrap
import time

def debugMsg(msg,type='Info'):
	out = ""
	out += type.upper() + ":" + "\n"
	out += "    " + msg + "\n" + "\n"

	print out

def checkArgs():
	#print('CLI arguments: ' + str(len(sys.argv)))
	#print()

	numArgs = len(sys.argv)
	if (numArgs < 1):
		#debugMsg("Please provide a path to the input file and the output file (number of CLI arguments: " + str(numArgs-1) + ". Expected: 2)",'Error')
		#print("Usage:")
		#print("   " + sys.argv[0] + " input output")
		#sys.exit(1)
		return False
	return True

# -----------------------------------------------------------------------------
# CONFIG

defaultInputFilePattern = "*"
#defaultInputFilePattern = "*.xlsx"
parameterValueSeparator = "="

try:
	filenameInput = sys.argv[1]
except IndexError:
	workingDir = os.getcwd()
	debugMsg('Looking for .' + defaultInputFilePattern + ' files in ' + workingDir)

	counter = 0
	matchingFiles = {}
	for file in os.listdir(workingDir):
		if fnmatch.fnmatch(file, defaultInputFilePattern):
			counter += 1
			matchingFiles[counter] = file
			print(str(counter) + ") " + file)

	if counter < 1:
		debugMsg('No ' + defaultInputFilePattern + ' files found in current directory. Please provide a path to the input file.','Info')
		filenameInput = raw_input("Input file (." + defaultInputFilePattern + "): ")
	else:
		print("")
		keyInput = int(raw_input("Input file number: "))

		if keyInput in matchingFiles:
			filenameInput = matchingFiles[keyInput]
		else:
			debugMsg('Option ' + str(keyInput) + ') does not exist ','Error')
			debugMsg('Please enter the filename by yourself','Info')
			filenameInput = raw_input("Input file (" + defaultInputFilePattern + "): ")

	#filenameOutput = raw_input("Output file (.xlsx): ")
	#print("")

if not os.path.exists(filenameInput):
	debugMsg("File '" + filenameInput + "' not found",'Error')
	criticalAbort()

filenameOutput = filenameInput + ".csv"

ptrInput = open(filenameInput,'rb')
ptrOutput = open(filenameOutput,'w')

section = ""
for line in ptrInput:
	if line[0] == "*":
		# Comment
		continue

	if len(line.strip()) < 1:
		continue;

	if re.search(':$',line):
		#print "Parameter: " + line
		section = line.replace(':','').strip()
		continue

	if re.search('[\s\t]+',line):
		#print "Value: " + line
		posEq = line.find(parameterValueSeparator)

		if posEq == -1:
			print "-- Something's fishy here: " + line
			continue

		cells = line.split(parameterValueSeparator)
		parameter = cells[0].strip()
		value = cells[1].strip()

		output = section + ";" + parameter + ";" + value + ";"

		print output
		ptrOutput.write(output + "\n")

	#print line

ptrInput.close()
ptrOutput.close()