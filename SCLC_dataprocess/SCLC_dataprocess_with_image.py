# For Caroline's data processing task
# Read multiple csv data files, calculate 2 columns according to raw data, combine together to 1 excel file.
#V-Vbi and SqrtJ

#Author: Charlie
# charlie_xumeng@hotmail.com
# 2022/03/19

#Updated on 7/4/2022 : Insert chart to the output file.

#import csv
import glob
import os
import pandas as pd


def main():

    # #set output file name
    OUTPUTFILE = "output.xlsx"
    # Data starting line number
    STARTLINE = 16
    # Used filename start
    NAMESTART = 7

    # #delete existed output file
    if os.path.exists(OUTPUTFILE):
        try:
            os.remove(OUTPUTFILE)
        except:
            input("Something went wrong. Try close applications using the output file.")
            return 0

    # read all the files in this folder and make a to do file list.
    filelist = glob.glob('*.csv')

    # ask for vbi value
    vbi = float(input("Input vbi value: "))

    # export to excel file
    with pd.ExcelWriter(OUTPUTFILE) as writer:
        for filenames in filelist:
            # use file name as worksheetname
            worksheetname = filenames[NAMESTART:]

            # build source dataframe (reading from csv file)
            source = pd.read_csv(filenames, usecols=[
                                 0, 2], skiprows=STARTLINE-1)

            # calculation new columns
            source["V-Vbi"] = source["Voltage (V)"] - vbi
            source["SqrtJ"] = source[" Current Density (mA/cm2)"]**(1/2)

            # write to excel with sheetname
            source.to_excel(writer, sheet_name=worksheetname)

            # Access the XlsxWriter workbook and worksheet objects from the dataframe.
            workbook = writer.book
            worksheet = writer.sheets[worksheetname]

            # create chart object
            chart = writer.book.add_chart({'type': 'scatter'})

            # get the number of rows and columns and length of source dataframe
            col_x = source.columns.get_loc('V-Vbi') + 1
            col_y = source.columns.get_loc('SqrtJ') + 1
            max_row = len(source)

            #set data in the chart
            chart.add_series({
                #'name':       "V-Vbi and SqrtJ",
                'categories': [worksheetname, 1, col_x, max_row, col_x],
                'values':     [worksheetname, 1, col_y, max_row, col_y],
                'marker':     {'type': 'circle', 'size': 4},
                # 'trendline': {'type': 'linear', 'display_equation': True,},
            })

            #set chart 
            chart.set_x_axis({'name':'V-Vbi'})
            chart.set_y_axis({'name':'SqrtJ', 'major_gridlines':{'visible':False}})
            chart.set_size({'width':800, 'height':600})
            chart.set_legend({'none': True})


            #insert chart to worksheet
            worksheet.insert_chart('G2',chart)


    print(str(len(filelist)) + " csv files are processed.")

    input("Check the result.")


if __name__ == '__main__':
    main()
