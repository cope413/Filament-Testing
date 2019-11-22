import os
from os import path


def import_settings(file_path):
	with open(file_path) as gcode_file:
		all_lines = gcode_file.readlines()

		# Filter to comments with equals

		comments = [comment for comment in all_lines if comment.startswith('; ') and len(comment.split('=')) == 2]
	# Extract settings from gcode comments into dictionary
	settings = {}

	for comment in comments:
		segments = [x.strip() for x in comment.replace(';', '').strip().split('=')]
		settings[segments[0]] = segments[1]

	return (settings, all_lines)


def import_mc_gcode():
	# MatterControl GCode directory
	directory = os.path.expandvars(r'\Users\taylo\Documents\Tensile Bar\Gcode')

	# Open output stream
	csv_stream = open('C:/Users/taylo/Downloads/tensile.csv', 'a')

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

	# Write headers
	# csv_stream.write('Name,')

	# for c in columns:
	# csv_stream.write("%s," % c)

	# csv_stream.write('\n')

	# Loop over all files in gcode folder
	for file_name in os.listdir(directory):

		file_name = file_name.lower()

		if file_name.endswith(".gcode"):
			# Collect the name without extension
			segments = os.path.splitext(file_name)
			name_without_extension = segments[0]

		print(name_without_extension)

		full_path = os.path.join(directory, file_name)

		# Extract settings
		(settings, all_lines) = import_settings(full_path)
		if len(settings) > 40:
			csv_stream.write("%s," % (name_without_extension))

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
		csv_stream.write('%s,' % nozzle_temp.strip())
		csv_stream.write('%s' % bed_temp.strip())
		csv_stream.write('\n')

	csv_stream.close()


# Kick off import on load
import_mc_gcode();

import shutil
source_path = "C:/Users/taylo/Documents/Tensile Bar/Gcode/"
source = os.listdir(source_path)
destination = "C:/Users/taylo/Documents/Tensile Bar/Old Files/"

for file_name in source:
	file_name = file_name.lower()
	if file_name.endswith(".gcode"):
		gcode = os.path.join(source_path, file_name)
		print(gcode)


def move_files():
	for files in source:
		if files.endswith(".gcode"):
			gcode = os.path.join(files)
			shutil.move(os.path.join(source_path, gcode), os.path.join(destination, gcode))


move_files();