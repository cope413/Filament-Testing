import csv
import os
import shutil
from mfi_google import login, upload

source_path = "C:/Python Projects/Filament Testing/MFI Results/"
destination_path = "C:/Python Projects/Filament Testing/Old MFI/"

def MFI_results(file_path):
	with open(file_path) as MFI_file:
		lines = MFI_file.readlines()
		comments = [comment for comment in lines]
	settings = {}

	for comment in comments:
		segments = [x.strip() for x in comment.strip().split(',')]
		settings[segments[0]] = segments[1]

	return settings, lines


def import_MFI_results():
	# MatterControl GCode directory
	directory = os.path.expandvars(r'\Python Projects\Filament Testing\MFI Results')
	mfiValues_batch = []
	# Open output stream
	csv_stream = open('C:/Python Projects/Filament Testing/mfi.csv', 'a')

	columns = ['"Material Reference"', '"Batch Reference"', '"Test Temperature (°C)"', '"Test Weight (kg)"',
	           '"Preheat Time (secs)"', '"Material Density (g/cc)"', '"No. of Tests"',
	           '"Batch Mean (g/10mins)"', '"Batch Std Dev."']

	# Loop over all files in folder
	for file_name in os.listdir(directory):

		file_name = file_name.lower()
		mfiValues = []
		if file_name.endswith(".csv"):
			# Collect the name without extension
			sections = os.path.splitext(file_name)
			name_without_extension = sections[0]

		full_path = os.path.join(directory, file_name)

		# Extract settings

		settings, lines = MFI_results(full_path)
		for c in columns:
			csv_stream.write("%s," % settings[c])
			mfiValues.append(settings[c])
		csv_stream.write('\n')
		mfiValues_batch.append(mfiValues)


	csv_stream.close()
	return mfiValues_batch



def move_files():
	mfi = os.listdir(source_path)
	for files in mfi:
		if files.endswith(".csv"):
			mfi_results = os.path.join(files)
			shutil.move(os.path.join(source_path, mfi_results), os.path.join(destination_path, mfi_results))
			print(mfi_results + " - file imported and moved successfully")


import_MFI_results()
mfiData = import_MFI_results()

if len(mfiData) > 0:
	creds = login()
	upload(creds, mfiData)
else:
	print(':: No MFI data')

move_files()
