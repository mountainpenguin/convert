#!/usr/bin/env python

import xlrd
import xlwt
import os
import re

"""
    A -> 0
    AA -> 26
    IP -> 249
"""
def conv(letters):
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    map = {
        "A" : "1",
        "B" : "2",
        "C" : "3",
        "D" : "4",
        "E" : "5",
        "F" : "6",
        "G" : "7",
        "H" : "8",
        "I" : "9",
        "J" : "A",
        "K" : "B",
        "L" : "C",
        "M" : "D",
        "N" : "E",
        "O" : "F",
        "P" : "G",
        "Q" : "H",
        "R" : "I",
        "S" : "J",
        "T" : "K",
        "U" : "L",
        "V" : "M",
        "W" : "N",
        "X" : "O",
        "Y" : "P",
        "Z" : "Q",
    }
    vals = ""
    for letter in letters:
        vals += map[letter.upper()]
    result = int(vals, 28) - 1
    result -= 2*(result / 27)
    if xlrd.colname(result) != letters.upper():
        print "Error cell reference out of range"
        return None
    return result
       
if __name__ == "__main__": 
    print "Excel ATPase results converter"
    print "------------------------------"
    inputfilename = raw_input("Input File: ")
    while True:
        if os.path.exists(inputfilename):
            break
        else:
            print "No such file, try again"
            inputfilename = raw_input("Input File: ")
    startcell = raw_input("Start Cell: ")
    endcell = raw_input("End Cell: ")
    outputfilename = raw_input("Output File: ")
    if "." not in outputfilename:
        outputfilename += ".xls"
    
    if inputfilename and startcell and endcell and outputfilename:
        try:
            startrow = int(re.findall("(\d+)", startcell)[0])
            startcol = conv(startcell.split(str(startrow))[0]) #convert letters to numbers
            endrow = int(re.findall("(\d+)", endcell)[0])
            endcol = conv(endcell.split(str(endrow))[0]) #convert letters to numbers
        except:
            print "Invalid cell references"
        else:
            if os.path.exists(inputfilename):
                try:
                    workbook = xlrd.open_workbook(inputfilename)
                    outputbook = xlwt.Workbook()
                    outputsheet = outputbook.add_sheet('Sheet 1')
                except:
                    print "Error, invalid file, try saving it as an 'Excel 97-2003 Workbook'"
                else:
                    newrownum = 1
                    for sheet in workbook.sheets():
                        sheetname = sheet.name
                        outputsheet.row(newrownum).write(0, sheetname)
                        newcolnum = 1
                        for rownum in range(startrow-1, endrow):
                            for colnum in range(startcol, endcol+1):
                                cellval = sheet.cell(rownum, colnum).value
                                outputsheet.row(newrownum).write(newcolnum, cellval)
                                newcolnum += 1
                        newrownum += 1
                    outputbook.save(outputfilename)
                    #create gnuplot information?
                    
            else:
                print "Error, no such file"
    else:
        print "Error, try again"