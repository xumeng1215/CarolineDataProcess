##For Caroline's data process
##Charlie
##Version 1
##2017-Jul-27

##update
##dark files is ignored.
import csv
import glob
import os

#set output file name
outputcsv = 'output.csv'
#outputcsv_Dark = 'output_Dark.csv'

#delete existed output file
if os.path.exists(outputcsv):
	os.remove(outputcsv)
#if os.path.exists(outputcsv_Dark):
#	os.remove(outputcsv_Dark)
	
#read all the files in this folder and make a to do file list.
filelist=glob.glob('*.csv')
#print (filelist)
print (str(len(filelist)) + ' files will be processed.')
#output csv header
head_list = ['Substrate:', 'Cell Number','Notes:','Cell Area (cm2):','Temperature (C):','Light Power (mW/cm2): ','Jsc (mA/cm2): ','Voc (V): ','FF: ','Eff (%): ','Rshunt (Ohms): ','Rseries (Ohms): ']
#head_list_Dark = ['Substrate:','Cell Number','Notes:','Cell Area (cm2):','Temperature (C):','Ideality [LnJ vs. V]: ','Jsat [Ln J vs. V] (A/m2): ','Jsat [Ln J vs. V] (mA/cm2): ','Schottky? : ','SBH [Ln J vs. V] (eV): ','R-Squared [Ln J vs. V]: ']
total_fields = ['Cell Name:', 'Notes:','Cell Area (cm2):','Temperature (C):','Light Power (mW/cm2): ','Jsc (mA/cm2): ','Voc (V): ','FF: ','Eff (%): ','Rshunt (Ohms): ','Rseries (Ohms): ','Cell Name:','Notes:','Cell Area (cm2):','Temperature (C):','Ideality [LnJ vs. V]: ','Jsat [Ln J vs. V] (A/m2): ','Jsat [Ln J vs. V] (mA/cm2): ','Schottky? : ','SBH [Ln J vs. V] (eV): ','R-Squared [Ln J vs. V]: ']

#write header
with open(outputcsv, 'a') as csvheader:
	head_writer=csv.writer(csvheader)
	head_writer.writerow(head_list)
#with open(outputcsv_Dark, 'a') as csvheader_Dark:
#	head_writer=csv.writer(csvheader_Dark)
#	head_writer.writerow(head_list_Dark)
	


#define a function to extract data and write into csv file as a line.
def write_into_csv(csvinput, csvoutput):
	lst=[]
	with open(csvinput, 'r') as inp:
		csvdata = csv.reader(inp)
		for line in csvdata:
			if line and (line[0] in total_fields):
					lst.append(line[1])
	
	#split the line[0] into 2 fields
	temp=lst[0].split('_')
	if len(temp) == 2:
		lst[0] = temp[0]
		lst.insert(1, temp[1])
	
	#print (lst)
	with open(csvoutput, 'a') as op:		
		datawriter = csv.writer(op)
		datawriter.writerow(lst)


#loop in the file list. call defined function.
for filename in filelist:
	if 'Dark.csv' not in filename:		
		write_into_csv(filename, outputcsv)
	#else:
	#	write_into_csv(filename, outputcsv)

