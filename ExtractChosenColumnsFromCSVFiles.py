import os
import csv
import xlwt

import tkinter as tk
from tkinter.filedialog import askdirectory

root = tk.Tk()
search_dir = askdirectory(title='Select folder with CSV files you want to columns extract from')

filename_subset = '.csv'
headers = ['field1', 'field2', 'field3', 'field4'] # headers to extract
print("Extracting ", headers, " from CSV files in ", search_dir, "\n") 
output = r'Output\Extraction.xls'


# Returns a list of the files that will be extracted from.
def GatherFiles(sdir, subset):
    lst = []
    for subdir, dirs, files in os.walk(sdir):
        for file in files:
            if subset in file:
                lst.append(os.path.join(subdir,file))
    return lst


# Input: CSV file
# Output: List of index values associated with the columns that will be extracted.
def LookUp(FileName):
    with open(FileName) as csvfile:
        content = csv.reader(csvfile, delimiter=',')
        for count, row in enumerate(content):
            if count == 0:
                csv_headers = row
            else:
                break
    return sorted(map(lambda x: csv_headers.index(x), list(set(filter(lambda x: x in headers, csv_headers)))))


# Main function that reads csv columns and writes the ones associated with the headers it to a new xls file
def Main():
    lst = GatherFiles(search_dir, filename_subset)
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Data")
    countc = 0
    for file in lst:
        Lookup_Index = LookUp(file)
        print(file)
        print("Columns Extracted:", Lookup_Index)
        print("")
        with open(file) as csvfile:
            read_content = csv.reader(csvfile, delimiter=',')
            for index in Lookup_Index:
                csvfile.seek(0) # moves back to the top of the csv
                for countr, row in enumerate(read_content):
                    ws.write(countr, countc, row[index])
                countc +=1 # increase column count
    wb.save(output)
    print("\nData extracted to:", output)

#Run the Main function             
Main()
root.destroy()
done = input('Press ENTER to close')
