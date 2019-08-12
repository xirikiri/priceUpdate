#!/usr/bin/env python
# -*- coding: utf8 -*-
import csv
import xlrd
import io
import os

def virtual_to_simple(file_obj):
    with open(file_obj, 'r') as m_file_exp, open('output.csv','w') as outp, open('missing_in_income.csv','w') as missing_in_income, open('kg.csv','w') as kg, open('sku_error.csv','w') as sku_error:
        reader = csv.reader(m_file_exp)
        #xls file
        book = xlrd.open_workbook("Excel.xls")
        print("The number of worksheets is {0}".format(book.nsheets))
        print("Worksheet name(s): {0}".format(book.sheet_names()))
        sh = book.sheet_by_index(0)
        print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
        print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))

        temp_import = []
        temp_missing = []
        temp_kg = []
        temp_sku_error = []

        for row in reader:
            temp_import.append(row)

        for row in temp_import:
            if row[27]=="no_selection":
                row[27]=row[25]
            if row[3]=="virtual":
                row[3] = "simple"
            flag = 0
            if (row[0]!="sku")&(row[3]!="configurable"):
                try:
                    for rx in range(sh.nrows):
                        if (float(row[0]) == sh.cell_value(rx,3)):
                            flag = 1
                            if (sh.cell_value(rx,5)=='кг'):
                                temp = ""
                                for i in range(0,len(row[6])):
                                    if (row[6][len(row[6])-i-1].isdigit()) or (row[6][len(row[6])-i-1]=='.'):
                                        temp += row[6][len(row[6])-i-1]
                                    if (row[6][len(row[6])-i-1]=='-'):
                                        break
                                weight = float(temp[::-1])
                                row[13] = sh.cell_value(rx,4)*weight
                                row[10] = 1
                            if (sh.cell_value(rx,5)=='шт'):
                                row[13]=sh.cell_value(rx,4)
                                row[10]=1
                            break
                except:
                    temp_sku_error.append(row)
                    flag = 1
                    row[10]=2
            if (flag == 0)&(row[3]!="configurable")&(row[0]!="sku"):
                temp_missing.append(row)
                row[10] = 2

        writer = csv.writer(outp)
        for row in temp_import:
            writer.writerow(row)
        writer = csv.writer(missing_in_income)
        for row in temp_missing:
            writer.writerow(row)
        writer = csv.writer(kg)
        for row in temp_kg:
            writer.writerow(row)
        writer = csv.writer(sku_error)
        for row in temp_sku_error:
            writer.writerow(row)

if __name__ == '__main__':
    export_file = "Export.csv"
    excel_file = "Excel.xls"
    virtual_to_simple(export_file)
    #Delete not needed files at current time
    if os.path.exists("missing_in_income.csv"):
    	os.remove("missing_in_income.csv")
    if os.path.exists("sku_error.csv"):
    	os.remove("sku_error.csv")
    if os.path.exists("kg.csv"):
    	os.remove("kg.csv")
