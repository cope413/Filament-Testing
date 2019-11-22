import csv
import os


with open('KVP PLA_Red_14932.csv', newline='') as csvfile:
	mfireader = csv.reader(csvfile)
	number = 3
	for row in mfireader:
		lines = csvfile.readlines()
		print(lines[2])
		print(number)
