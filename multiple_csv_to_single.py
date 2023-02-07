import openpyxl
import os
import csv

# set the directory where the cs files are located
dirname = '.'

# create a list to store the filenames
filenames = []

# loop through the directory and add each filename to the list
for filename in os.listdir (dirname):
    if os.path.splitext(filename) [1] == " csv":
     filenames.append(filename)

# open the output file
wb = openpyxl.Workbook()

# loop through each file and write its contents to the output file
for filename in filenames:
    with open(os.path. join(dirname, filename), 'r') as infile:
        W5 = wb. create_sheet(filename)
        reader = csv.reader(infile)
        for row in reader:
            W5.append(row) 

# save the output file
wb.save ('combine_files.xlsx')


