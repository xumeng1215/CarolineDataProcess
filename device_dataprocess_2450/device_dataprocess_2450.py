# Build by Charlie XU and Caroline Pan
# 30/03/2023
# charlie_xumeng@hotmail.com
# Keithley sourcemeter keeps generating CSV files in a folder as testing results.
# This script process existing CSV files, and keeps monitoring this folder to process new generated files.
# While processing, calculated numbers (Isc, Jsc, Voc, FF etc.) will be output to console
# And to a different CSV file for record as well.

import csv
import os
import time

# define useful data area in the testing result csv file
VOLTAGE_COLUMN_INDEX = 2
CURRENT_COLUMN_INDEX = 3
POWER_COLUMN_INDEX = 4
START_ROW = 35
#FINISH_ROW = 134

# result file for the record
result_file = "device_result_" + time.strftime("%Y-%m-%d_%H-%M", time.localtime()) + ".csv"

# Input the folder path of testing result csv files, the folder should be created already
folder_path = input("Enter the folder path of output csv files: ")

# Check if the folder exists
if not os.path.isdir(folder_path):
    print("Invalid folder path.")
    exit()

# Input area value of the device for testing
area = float(input("Please enter the area of the device in cm^2: "))

# Give default input power density value, this value is decided by the light source.
print("Start with input power density as 100mW/cm^2. Press Ctrl-C to stop the monitoring. ")
pin = float(100)

# define a function to extract the csv file to a table
# def extract_columns(csv_file, col_nums, start_row, end_row):
#     with open(csv_file, 'r') as file:
#         reader = csv.reader(file)
#         table = []
#         for i, row in enumerate(reader):
#             if i >= start_row and i <= end_row:
#                 table_row = []
#                 for col_num in col_nums:
#                     table_row.append(row[col_num])
#                 table.append(table_row)
#     return table

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

# define a function to return the position of 2 numbers next to zero in a list, not the actual value of the number
# the list will have a number series goes from negative to positive
# so the two closest to zero numbers should be one negative and one positive


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


# define a funtion to process the csv file, output the value in console and result csv file.
# return a dictionary which have value calculated from this input csv file
# including Isc(), Voc(), Pmax(), Jsc(), FF(), Pce()
def process(csv_file):
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

    # output the values in console
    print("{:20} {:20.4f} {:20.4f} {:20.4f} {:20.4f} {:20.4f}".format(csv_file,
                                                                      jsc, voc, pmax, ff, pce))

    # output the values to result file
    with open(result_file, 'a', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow([csv_file, csv_file.split(' ')[0], jsc, voc, pmax, ff, pce])
    # return a dictionary
    return {"isc": isc, "voc": voc, "pmax": pmax, "jsc": jsc, "pce": pce, "ff": ff}


# Initialize the latest file name and timestamp
latest_file = None
latest_timestamp = 0

# Initialize a flag to track whether the latest file has been output
latest_file_outputted = False


# Output the header of all the data columns: JSC VOC PMAX FF PCE
# to console
print("{:>20} {:>20} {:>20} {:>20} {:>20} {:>20}".format("Device Name",
                                                         "Jsc (mA/cm^2)", "Voc (V)", "Pmax (mW/cm^2)", "FF (%)", "PCE (%)"))

# and to result file for record
with open(result_file, 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["CSV File", "Cell Code", "Jsc (mA/cm^2)", "Voc (V)",
                    "Pmax (mW/cm^2)", "FF (%)", "PCE (%)"])


# Process exsiting files
files = os.listdir(folder_path)
for file in files:
    file_fullpath = os.path.join(folder_path, file)
    timestamp = os.path.getmtime(file_fullpath)
    if timestamp > latest_timestamp:
        # Update the latest file name and timestamp
        latest_file = file
        latest_timestamp = timestamp
    dict = process(file)
    latest_file_outputted = True

# Monitor the folder for new files
while True:
    # Get a list of all files in the folder
    files = os.listdir(folder_path)

    # Loop through each file to the latest one
    for file in files:
        # Get the full path of the file
        file_path = os.path.join(folder_path, file)

        # Get the timestamp of the file
        timestamp = os.path.getmtime(file_path)

        # Check if the file is newer than the latest file
        if timestamp > latest_timestamp:
            # Update the latest file name and timestamp
            latest_file = file
            latest_timestamp = timestamp
            # Reset the flag indicating whether the latest file has been output
            latest_file_outputted = False

    # Check if a new file has been found and not outputted yet
    if latest_file is not None and not latest_file_outputted:
        latest_filepath = os.path.join(folder_path, latest_file)
        print("Processing csv file: " + latest_file)

        # use returned dictionary
        dict = process(latest_file)

        # Set the flag indicating that the latest file has been output
        latest_file_outputted = True

    # Wait for 1 second before checking again
    time.sleep(1)
