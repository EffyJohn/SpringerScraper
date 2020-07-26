from bs4 import BeautifulSoup
import xlrd
import urllib.request
from os import system
import csv

springer_excel = xlrd.open_workbook("./Springer_books.xlsx")
sheet = springer_excel.sheet_by_index(0)

nrows = sheet.nrows
ncols = sheet.ncols

print(str(ncols) + " " + str(nrows))

table = []

index = 118

while (index < nrows):
    current_row = []

    for j in range(ncols):
        current_row.append(sheet.cell_value(index,j))

    table.append(current_row)
    index += 1

flag = False

new_table = []

iterable = 0

try:
    for element in table:
        print(str(iterable) + "/291 files completed. " + str(round(iterable/2.91, 2)) + "% completed")

        if flag == False:
            flag = True
            new_table.append(element)
            continue
        
        print("Accessing HTML")
        document = urllib.request.urlopen(element[4])

        print("Parsing Downloaded HTML")
        soup = BeautifulSoup(document, 'html.parser')

        element[5] = ("link.springer.com" + soup.find_all("a", class_="test-bookpdf-link")[0]['href'])

        print("Adding Link to list")
        new_table.append(element)
        system('clear')
        iterable += 1
except:
    print("Fatal Error: Writing CSV and quitting")
    with open("output3.csv", "w", newline="") as out_file:
        writer = csv.writer(out_file)
        writer.writerows(new_table)
    exit()

print("Success! Writing CSV and quitting")
with open("output3.csv", "w", newline="") as out_file:
        writer = csv.writer(out_file)
        writer.writerows(new_table)