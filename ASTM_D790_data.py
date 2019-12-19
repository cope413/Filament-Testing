
# THIS IS FOR ASTM D790 TEST REPORTS ONLY!

import os
import pandas as pd
import csv
import shutil
from ASTM_D790_google import login, upload
data_source = "C:/Python Projects/Filament Testing/Flexural Data/ASTM D790/"
xls_destination = "C:/Python Projects/Filament Testing/Old Flexural Data/Excel/ASTM D790/"
csv_destination = "C:/Python Projects/Filament Testing/Old Flexural Data/CSV/ASTM D790/"
output = "C:/Python Projects/Filament Testing/ASTM D790.csv"


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


def import_Flexural_data():
	tensile_directory = os.path.expandvars(data_source)
	flexuralValuesBatch = []
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
					test_operator = line[11]
					material = line[9]
					batch_number = line[16]
					# print("Test Report - " + report_number)
					break
				for row in reader:   # skips 2nd row of Tensile data report
					for row in reader: # writes data fields of Tensile reports to csv file
						flexuralValues = []
						specimen_width = row[12]
						specimen_thickness = row[13]
						max_Force_lbs= row[14]
						Peak_Flexural_Stress_psi = row[15]
						Flexural_Strain_break_inin = row[16]
						Modulus_Elasticity_psi = row[17]
						test_number = row[10]
						specimen_ID = row[11]
						comments = row[18]
						writeFile.write("%s," % specimen_ID)
						writeFile.write("%s," % test_number)
						writeFile.write("%s," % report_number)
						writeFile.write("%s," % test_date)
						writeFile.write("%s," % test_operator)
						writeFile.write("%s," % material)
						writeFile.write("%s," % batch_number)
						writeFile.write("%s," % specimen_width)
						writeFile.write("%s," % specimen_thickness)
						writeFile.write("%s," % max_Force_lbs)
						writeFile.write("%s," % Peak_Flexural_Stress_psi)
						writeFile.write("%s," % Flexural_Strain_break_inin)
						writeFile.write("%s," % Modulus_Elasticity_psi)
						writeFile.write("%s," % comments)
						flexuralValues.append(specimen_ID)
						flexuralValues.append(test_number)
						flexuralValues.append(report_number)
						flexuralValues.append(test_date)
						flexuralValues.append(test_operator)
						flexuralValues.append(material)
						flexuralValues.append(batch_number)
						flexuralValues.append(specimen_width)
						flexuralValues.append(specimen_thickness)
						flexuralValues.append(max_Force_lbs)
						flexuralValues.append(Peak_Flexural_Stress_psi)
						flexuralValues.append(Flexural_Strain_break_inin)
						flexuralValues.append(Modulus_Elasticity_psi)
						flexuralValues.append(comments)
						writeFile.write('\n')
						flexuralValuesBatch.append(flexuralValues)

	return flexuralValuesBatch


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
import_Flexural_data()

flexuralData = import_Flexural_data()

if len(flexuralData) > 0:
	creds = login()
	upload(creds, flexuralData)
else:
	print(':: No data')

move_CSV_files()