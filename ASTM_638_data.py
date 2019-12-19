
# THIS IS FOR ASTM 638 TEST REPORTS ONLY!

import os
import pandas as pd
import csv
import shutil
from ASTM_638_google import login, upload
data_source = "C:/Python Projects/Filament Testing/Tensile Data/ASTM 638/"
xls_destination = "C:/Python Projects/Filament Testing/Old Tensile Data/Excel/ASTM 638/"
csv_destination = "C:/Python Projects/Filament Testing/Old Tensile Data/CSV/ASTM 638/"
output = "C:/Python Projects/Filament Testing/ASTM 638.csv"


def convert_XLS():
	directory = os.path.expandvars(data_source)
	print(':: Converting XLS to CSV ::')
	for file_name in os.listdir(directory):
		file_name = file_name.lower()
		if file_name.endswith(".xls"):
			split = os.path.splitext(file_name)
			xls_name = split[0]
			read_file = pd.read_excel(os.path.join(data_source, file_name))
			read_file.to_csv(os.path.join(data_source, '%s.csv' % xls_name), index=None, header=True)
			shutil.move(os.path.join(data_source, file_name), os.path.join(xls_destination, file_name))


def import_Tensile_data():
	tensile_directory = os.path.expandvars(data_source)
	tensileValuesBatch = []
	for file_name in os.listdir(tensile_directory):
		if file_name.endswith(".csv"):
			chunks = os.path.splitext(file_name)
			csvName = chunks[0]
		full_path = os.path.join(tensile_directory, file_name)

		with open(full_path, 'r') as readFile:
			with open(output, 'a') as writeFile:
				reader = csv.reader(readFile)

				for line in reader:  # stores Tensile Report Batch data fields
					report_number = line[3]
					test_date = line[5]
					test_operator = line[9]
					material = line[12]
					batch_number = line[15]
					# print("Test Report - " + report_number)
					break
				for row in reader:   # skips 2nd row of Tensile data report
					for row in reader: # writes data fields of Tensile reports to csv file
						tensileValues = []
						test_number = row[10]
						specimen_ID = row[11]
						yield_LBS = row[12]
						yield_Strength_psi = row[13]
						tensile_LBS = row[14]
						modulus_Elasticity_psi = row[15]
						break_Elongation_in = row[16]
						yield_Elongation_in = row[17]
						tensile_Elongation_in = row[18]
						tensile_Strength_psi = row[19]
						comments = row[20]
						writeFile.write("%s," % specimen_ID)
						writeFile.write("%s," % test_number)
						writeFile.write("%s," % report_number)
						writeFile.write("%s," % test_date)
						writeFile.write("%s," % test_operator)
						writeFile.write("%s," % material)
						writeFile.write("%s," % batch_number)
						writeFile.write("%s," % yield_LBS)
						writeFile.write("%s," % yield_Strength_psi)
						writeFile.write("%s," % tensile_LBS)
						writeFile.write("%s," % modulus_Elasticity_psi)
						writeFile.write("%s," % break_Elongation_in)
						writeFile.write("%s," % yield_Elongation_in)
						writeFile.write("%s," % tensile_Elongation_in)
						writeFile.write("%s," % tensile_Strength_psi)
						writeFile.write("%s," % comments)
						tensileValues.append(specimen_ID)
						tensileValues.append(test_number)
						tensileValues.append(report_number)
						tensileValues.append(test_date)
						tensileValues.append(test_operator)
						tensileValues.append(material)
						tensileValues.append(batch_number)
						tensileValues.append(yield_LBS)
						tensileValues.append(yield_Strength_psi)
						tensileValues.append(tensile_LBS)
						tensileValues.append(modulus_Elasticity_psi)
						tensileValues.append(break_Elongation_in)
						tensileValues.append(yield_Elongation_in)
						tensileValues.append(tensile_Elongation_in)
						tensileValues.append(tensile_Strength_psi)
						tensileValues.append(comments)
						writeFile.write('\n')
						tensileValuesBatch.append(tensileValues)

	return tensileValuesBatch


def move_CSV_files():
	directory = os.path.expandvars(data_source)
	print(':: Moving CSV files ::')
	for file_name in os.listdir(directory):
		file_name = file_name.lower()
		if file_name.endswith(".csv"):
			split = os.path.splitext(file_name)
			csv_name = split[0]
			print("Successfully imported & moved:: " + csv_name + ".csv")
			shutil.move(os.path.join(directory, file_name), os.path.join(csv_destination, file_name))


convert_XLS()
import_Tensile_data()

tensileData = import_Tensile_data()

if len(tensileData) > 0:
	creds = login()
	upload(creds, tensileData)
else:
	print(':: No data')

move_CSV_files()