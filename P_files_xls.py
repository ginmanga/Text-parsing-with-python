"""Program to extract data from excel files."""
# ------------------------------------------------------------------------------
# Name:        Extract Data Excel for Kay
# Purpose:     Extract Data Excel for Kay
#
# Author:      Jorge Patron Boenheim
#
# Created:     17/02/2016
# Copyright:   (c) Jorge 2016
# Licence:     <your licence>
# ------------------------------------------------------------------------------
import xlrd
import os
import csv
import re
RE = re.compile(u'[⺀-⺙⺛-⻳⼀-⿕々〇〡-〩〸-〺〻㐀-䶵一-鿃豈-鶴侮-頻並-龎]', re.UNICODE)
directory_a = os.getcwd()
interest_items = []

x = int(input("Hello!, please specify the column number with the identifying"
              " code:"))
x = x-1
temp_inp = input("Please specify column numbers to extract separated"
                 " by a comma:")
# Reduce column number to match python list indexing
yy = [int(i) for i in temp_inp.split(',')]
y = []
[y.append(i-1) for i in yy]


def remove(x):
    """Function: Strip Chinese characters and lt spaces."""
    x = x.strip()
    x = RE.sub('', x)
    x = x.strip()
    return x


def clear_idcode(x):
    """Function to select code when there is an unwanted character."""
    return x


# Main loop
for i in os.listdir(directory_a):
    ii = directory_a+"/"+i
# Try to open file using xlrd.open_workbook function
# If it fails, then do not perform operations on that file
    try:
        c_file = xlrd.open_workbook(filename=ii)
        cc = "opening file: "+ii
        print(cc)
    except:
        continue
# Create sheet object from the book object
    sheet_list = c_file.sheets()
    for sheet in c_file.sheets():
        row_counter = 0
        row_numbers = []
        for r in sheet.col(x):
            if r.ctype != 0 and r.value != "品种号" and r.value != "老":
                row_numbers.append(row_counter)
            row_counter += 1
        for row_n in row_numbers:
            a = sheet.cell(row_n, x)
            if a.ctype == 2:
                aa = str(int(a.value))
                b = [i, a.value, aa]
                b.extend([sheet.cell(row_n, xx).value for xx in y])
            else:
                aa = remove(a.value)
                # Fix multiple numbers separated by slashes
                # Can turn into a function for better code and functionality
                fslash = aa.find('/')
                fdot = aa.find('.')
                if fdot > -1:
                    print(aa)
                    print(type(aa))
                if fslash > -1:
                    aa1 = aa[0:fslash]
                    aa2 = aa[fslash+1:len(aa)]
                    fslash2 = aa2.find('/')
                    try:
                        int(aa1)
                        go = 1
                    except:
                        go = 0
                    try:
                        int(aa2)
                        go1 = 1
                    except:
                        go1 = 0
                    if fslash2 == -1 and go == 1 and go1 == 1:
                        if int(aa1) > int(aa2):
                            aa = aa1
                        else:
                            aa = aa2
                    if fslash2 == -1 and go == 0 and go1 == 1:
                        aa = aa2
                    if fslash2 == -1 and go == 1 and go1 == 0:
                        aa = aa1
                b = [i, a.value.strip(), aa]
                b.extend([sheet.cell(row_n, xx).value for xx in y])
            interest_items.append(b)
# Operation done only for the first sheet in each book.
# break command finishes loop on first sheet
        break

print("Printing to file")
file_name = directory_a + "/" + "output.csv"

with open(file_name, 'w', newline='', encoding='utf-8-sig') as csvfile:
    owrite = csv.writer(csvfile)
    owrite.writerows(interest_items)
