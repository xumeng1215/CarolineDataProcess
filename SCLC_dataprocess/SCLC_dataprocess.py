#For Caroline's data processing task
#Read multiple csv data files, calculate 2 columns according to raw data, combine together to 1 excel file.

#Author: Charlie
#charlie_xumeng@hotmail.com
#2022/03/19

import csv
import glob
import os
import pandas as pd


def main():
    
    # #set output file name
    OUTPUTFILE = "output.xlsx"
    #Data starting line number
    STARTLINE = 16
    #Used filename start
    NAMESTART = 14
    
    # #delete existed output file
    if os.path.exists(OUTPUTFILE):
        os.remove(OUTPUTFILE)
    
    #read all the files in this folder and make a to do file list.
    filelist=glob.glob('*.csv')   
    
    #ask for vbi value
    vbi = float(input("Input vbi value: "))


    #export to excel file
    with pd.ExcelWriter(OUTPUTFILE) as writer: 
        for names in filelist:        
            source= pd.read_csv(names, usecols=[0,2], skiprows=STARTLINE-1)
            source["V-Vbi"] = source["Voltage (V)"] - vbi
            source["SqrtJ"] = source[" Current Density (mA/cm2)"]**(1/2)
            source.to_excel(writer, sheet_name=names[NAMESTART:])

    print(str(len(filelist))+ " csv files are processed.")

    input("Check the result.")
        
if __name__ == '__main__':
    main()


