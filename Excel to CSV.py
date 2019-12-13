import os
import pandas as pd
import csv
import shutil

xls_source = "C:/Users/taylo/Documents/Filament-Testing/Tensile Data/"
xls_destination = "C:/Users/taylo/Documents/Filament-Testing/Old Tensile Data/Excel/"
csvFile = "C:/Users/taylo/Documents/Filament-Testing/Tensile Data/test report - 138.csv"
output = "C:/Users/taylo/Documents/Filament-Testing/Tensile Data/tensile.csv"

def convert_XLS():
	directory = os.path.expandvars(xls_source)
	print(':: Converting XLS to CSV ::')
	for file_name in os.listdir(directory):
		file_name = file_name.lower()
		if file_name.endswith(".xls"):
			split = os.path.splitext(file_name)
			xls_name = split[0]
			print(xls_name)
			read_file = pd.read_excel(os.path.join(xls_source, file_name))
			read_file.to_csv(os.path.join(xls_source, '%s.csv' % xls_name), index=None, header=True)
			shutil.move(os.path.join(xls_source, file_name), os.path.join(xls_destination, file_name))


with open(csvFile, 'r') as readFile:
	with open(output, 'a') as writeFile:
		reader = csv.reader(readFile)
		for line in reader:
			report_number = line[3]
			test_date = line[5]
			test_operator = line[9]
			material = line[12]
			batch_number = line[15]
			print(report_number)
			for row in reader:
				test_number = row[10]
				specimen_ID = row[11]
				yield_N = row[12]
				yield_Strength = row[13]
				tensile_N = row[14]
				modulus_Elasticity = row[15]
				break_Elongation = row[16]
				yield_Elongation = row[17]
				tensile_Elongation = row[18]
				tensile_Strength = row[19]
				comments = row[20]
				writeFile.write("%s," % specimen_ID)
				writeFile.write("%s," % test_number)
				writeFile.write("%s," % report_number)
				writeFile.write("%s," % test_date)
				writeFile.write("%s," % test_operator)
				writeFile.write("%s," % material)
				writeFile.write("%s," % batch_number)
				writeFile.write("%s," % yield_N)
				writeFile.write("%s," % yield_Strength)
				writeFile.write("%s," % tensile_N)
				writeFile.write("%s," % modulus_Elasticity)
				writeFile.write("%s," % break_Elongation)
				writeFile.write("%s," % yield_Elongation)
				writeFile.write("%s," % tensile_Elongation)
				writeFile.write("%s," % tensile_Strength)
				writeFile.write("%s," % comments)
				writeFile.write('\n')


"""with open(csvFile, 'r') as readFile:
	with open(output, 'w') as writeFile:
		reader = csv.DictReader(readFile, fieldnames=['TESTNUM', 'SPEC_ID', 'RESULT1', 'RESULT2', 'RESULT3', 'RESULT4',
		                                              'RESULT5', 'RESULT6', 'RESULT7', 'RESULT8', 'COMMENTS'])
		for row in reader:
			print(row['TESTNUM'])
"""

#convert_XLS()
#Tensile_Data()
