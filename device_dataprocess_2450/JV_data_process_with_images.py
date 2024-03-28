# Author: Charlie
# charlie_xumeng@hotmail.com
# 2023/11/07


# import csv
import glob
import os
import pandas as pd

# #set output file name
OUTPUTFILE = "output.xlsx"
    
VOLTAGE_COLUMN_INDEX = 2
CURRENT_COLUMN_INDEX = 3
# Data starting line number
STARTLINE = 35
FOOTROWS = 6

#find 2 closest numbers to zero in a list of numbers
#return 2 index number
def find_two_closest_to_zero(nums):
    closest_index = None
    second_closest_index = None
    for i, num in enumerate(nums):
        if closest_index is None or abs(num) < abs(nums[closest_index]):
            second_closest_index = closest_index
            closest_index = i
        elif second_closest_index is None or abs(num) < abs(nums[second_closest_index]):
            second_closest_index = i
    return closest_index, second_closest_index
 
#extract columns from a csv file, return as a table
def extract_columns(csv_file, col_nums, start_row):
    with open(csv_file, 'r') as file:
        reader = csv.reader(file)
        table = []
        for i, row in enumerate(reader):
            if i >= start_row:
                if not any(row):
                    break
                table_row = []
                for col_num in col_nums:
                    table_row.append(row[col_num])
                table.append(table_row)
    return table
 

def calculate_CSV(csv_file):
    csv_file_fullpath = os.path.join(folder_path, csv_file)
    # get 3 columns for voltage and current and power
    columns = extract_columns(csv_file_fullpath, [
                              VOLTAGE_COLUMN_INDEX, CURRENT_COLUMN_INDEX, POWER_COLUMN_INDEX], START_ROW)
    v_column = []
    i_column = []
    p_column = []
    for row in columns:
        v_column.append(float(row[0]))
        i_column.append(float(row[1]))
        p_column.append(float(row[2]))

    
    # assuming linear relationship I = kV + A 
    # calculate the isc (estimated value in i_column when v=0)
    #
    i, j = find_two_closest_to_zero(v_column)    
    k1 = (i_column[i] - i_column[j])/(v_column[i] - v_column[j])
    isc = i_column[j] - k1 * v_column[j]

    # calculate the voc (estimated value in v_column when i=0)
    #
    m, n = find_two_closest_to_zero(i_column)
    k2 = (i_column[m]-i_column[n])/(v_column[m]-v_column[n])
    voc = v_column[n] - i_column[n]/k2

    # calculate pmax in mW/cm^2, pmax should be negative
    #
    pmax = abs(min(p_column))*1000/area

    # calculate Jsc in mA/cm^2
    #
    jsc = abs(isc) * 1000 / area

    # calculate pce
    #
    pce = pmax * 100 / pin

    # calculate ff
    #
    ff = pmax * 100 / (jsc*voc)

    # output the values to result file
    # with open(result_file, 'a', newline='') as csvfile:
        # writer = csv.writer(csvfile)
        # writer.writerow([csv_file, jsc, voc, pmax, ff, pce])
    
    # return a dictionary
    return {"isc": isc, "voc": voc, "pmax": pmax, "jsc": jsc, "pce": pce, "ff": ff}

def main():
    
    # #delete existed output file
    if os.path.exists(OUTPUTFILE):
        try:
            os.remove(OUTPUTFILE)
        except:
            input("Something went wrong. Try close applications using the output file.")
            return 0

    # read all the files in this folder and make a to do file list.
    filelist = glob.glob("*.csv")

    # export to excel file
    with pd.ExcelWriter(OUTPUTFILE) as writer:
        for filenames in filelist:
            # use file name as worksheetname
            worksheetname = filenames

            # build source dataframe (reading from csv file)
            # read column 2 as Voltage, column 3 as Current,
            # starting at STARTLINE-1 to include the header, skipfooter 6 rows to get rid of the unneccessary values, otherwise type will change to string.

            source = pd.read_csv(
                filenames,
                usecols=[VOLTAGE_COLUMN_INDEX, CURRENT_COLUMN_INDEX],
                skiprows=STARTLINE - 1,
                skipfooter=FOOTROWS,
                engine="python",
            )

            # calculation new columns
            # the area of test device is 0.1 cm2
            source["Current density(mA/cm2)"] = source["SMU-1 Current (A)"] / 0.1 * 1000

            # write to excel with sheetname
            source.to_excel(writer, sheet_name=worksheetname)

            # Access the XlsxWriter workbook and worksheet objects from the dataframe.
            workbook = writer.book
            worksheet = writer.sheets[worksheetname]

            # print(source)
            
            # create chart object
            chart = writer.book.add_chart({"type": "scatter"})

            # get the number of rows and columns and length of source dataframe
            # +1 to get rid of the headers
            col_x = source.columns.get_loc("SMU-1 Voltage (V)") + 1
            col_y = source.columns.get_loc("Current density(mA/cm2)") + 1
            max_row = len(source)

            # set data in the chart
            chart.add_series(
                {
                    "categories": [worksheetname, 1, col_x, max_row, col_x],
                    "values": [worksheetname, 1, col_y, max_row, col_y],
                    "marker": {"type": "circle", "size": 4},
                }
            )

            # set chart
            chart.set_x_axis({"name": "SMU-1 Voltage (V)"})
            chart.set_y_axis(
                {
                    "name": "Current density (mA/cm2)",
                    "major_gridlines": {"visible": False},
                }
            )
            chart.set_size({"width": 800, "height": 600})
            chart.set_legend({"none": True})

            # insert chart to worksheet
            worksheet.insert_chart("J4", chart)

    print(str(len(filelist)) + " csv files are processed.")

    input("Check the result.")


if __name__ == "__main__":
    main()
