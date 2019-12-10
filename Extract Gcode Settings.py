#!/usr/bin/env python3

import os
from os import path
import shutil
from google_sheets import login, upload
import hashlib

print(os.path.realpath(__file__))
print(__file__)
source_path = "New Gcode/"
destination_path = "Old Gcode/"
output_csv = "slice-settings.csv"
if os.path.isdir('H:') is True:
	usbSource_path = "H:"
else:
	usbSource_path = None
usbSource = os.listdir(usbSource_path)

for file_name in usbSource:
	file_name = file_name.lower()
	if file_name.endswith(".gcode"):
		gcode = os.path.join(usbSource_path, file_name)
		print(gcode + " - USB file")


def USB_gcode():
	print("Moving files from USB stick")
	for files in usbSource:
		if files.endswith(".gcode"):
			gcode = os.path.join(files)
			shutil.move(os.path.join(usbSource_path, gcode), os.path.join(source_path, gcode))


'''
def date():
	from datetime import date
	now = date.today()
	newBatch = open('C:/Users/taylo/Downloads/tensile.csv', 'a')
	newBatch.write('%s, \n' % now)
	newBatch.close
'''


def import_settings(file_path):
	with open(file_path) as gcode_file:
		all_lines = gcode_file.readlines()
		gcodeString = ''.join([str(specimenID) for specimenID in all_lines])
		ID = hashlib.sha1(gcodeString.encode())

		# Filter to comments with equals

		comments = [comment for comment in all_lines if comment.startswith('; ') and len(comment.split('=')) == 2]
	# Extract settings from gcode comments into dictionary
	settings = {}

	for comment in comments:
		segments = [x.strip() for x in comment.replace(';', '').strip().split('=')]
		settings[segments[0]] = segments[1]

	return {'settings': settings, 'all_lines': all_lines, 'id': ID.hexdigest()}


def import_mc_gcode():
	print(':: Importing data from GCode files')

	# MatterControl GCode directory
	directory = os.path.expandvars(source_path)

	# Open output stream
	csv_stream = open(output_csv, 'a')

	columns = ['numberOfBottomLayers', 'numberOfPerimeters', 'numberOfTopLayers', 'outsidePerimeterExtrusionWidth',
	           'outsidePerimeterSpeed', 'firstLayerSpeed', 'topInfillSpeed', 'firstLayerExtrusionWidth',
	           'firstLayerThickness', 'minimumTravelToCauseRetraction', 'retractionOnTravel', 'retractionZHop',
	           'unretractExtraExtrusion', 'retractRestartExtraTimeToApply', 'retractionSpeed', 'bridgeSpeed',
	           'airGapSpeed', 'bottomInfillSpeed', 'bridgeOverInfill', 'extrusionMultiplier', 'infillStartingAngle',
	           'infillExtendIntoPerimeter', 'infillSpeed', 'infillType', 'minimumExtrusionBeforeRetraction',
	           'minimumPrintingSpeed', 'insidePerimetersSpeed', 'fanSpeedMinPercent', 'coastAtEndDistance',
	           'minFanSpeedLayerTime', 'fanSpeedMaxPercent', 'maxFanSpeedLayerTime', 'bridgeFanSpeedPercent',
	           'firstLayerToAllowFan', 'minimumLayerTimeSeconds', 'travelSpeed', 'filamentDiameter', 'layerThickness',
	           'extrusionWidth', 'avoidCrossingPerimeters', 'outsidePerimetersFirst', 'retractWhenChangingIslands',
	           'expandThinWalls', 'MergeOverlappingLines', 'fillThinGaps', 'infillPercent',
	           'perimeterStartEndOverlapRatio', 'filament used']

	# Write headers if file is empty
	if os.stat(output_csv).st_size == 0:
		csv_stream.write('Name,')
		for c in columns:
			csv_stream.write("%s," % c)

	csv_stream.write('\n')

	values_batch = []

	# Loop over all files in gcode folder
	for file_name in os.listdir(directory):

		if file_name.endswith(".gcode"):
			# Collect the name without extension
			segments = os.path.splitext(file_name)
			name_without_extension = segments[0]

		print(name_without_extension)

		full_path = os.path.join(directory, file_name)

		values = []

		# Extract settings
		dict = import_settings(full_path)

		settings = dict['settings']
		all_lines = dict['all_lines']
		ID = dict['id'].upper()

		if len(settings) > 40:
			csv_stream.write("%s," % (name_without_extension))
			values.append(name_without_extension)

		temp_line = [l for l in all_lines if l.startswith("M109")][0]
		bed_line = [x for x in all_lines if x.startswith("M190")][0]

		segments = temp_line.split(' ')
		segments1 = bed_line.split(' ')

		if len(segments) == 2:
			nozzle_temp = segments[1].replace('S', '')
		else:
			nozzle_temp = segments[2].replace('S', '')

		if len(segments1) == 1:
			bed_temp = segments1[0].replace('S', '')
		else:
			bed_temp = segments1[1].replace('S', '')

		for c in columns:
			csv_stream.write("%s," % settings[c])
			values.append(settings[c])

		csv_stream.write('%s,' % nozzle_temp.strip())
		values.append(nozzle_temp.strip())
		csv_stream.write('%s' % bed_temp.strip())
		values.append(bed_temp.strip())
		csv_stream.write('%s' % ID[0:6])
		values.append(ID[0:6])

		csv_stream.write('\n')
		values_batch.append(values)

	csv_stream.close()
	return values_batch


def move_files():
	print(':: Moving files')
	source = os.listdir(source_path)
	for files in source:
		if files.endswith(".gcode"):
			gcode = os.path.join(files)
			shutil.move(os.path.join(source_path, gcode), os.path.join(destination_path, gcode))


# Kick off import on load
USB_gcode()
data = import_mc_gcode()

if len(data) > 0:
	creds = login()
	upload(creds, data)
else:
	print(':: No data')

move_files()
