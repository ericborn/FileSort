# -*- coding: utf-8 -*-
"""
=============================================
Author:      Eric Born
Create date: 15 Dec 2018
Description: Used to create folders and sort files by reading data in an excel document
CS521 Final project

Last modified date:

Change log:

=============================================
"""
import pandas as pd
import os
import shutil

#import sys
#sys.path

#Imports class
from Input import Input

def main():
    #Creates a check that the file has a proper extension, xlsx or xls
    while True:
        #Calls setPath method to collect the filepath from the user
        inputPath = Input.setPath()
        if inputPath[-5:] == '.xlsx':
            #Finds the last double slash, indicating the path minus the actual file itself
            path = inputPath.rfind('\\')
            #sets the path and drops the file name and extension
            fullPath = inputPath[:path]
            #change directory so the file can be parsed by pandas
            os.chdir(fullPath)
            break
        elif inputPath[-4:] == '.xls':
            #Finds the last double slash, indicating the path minus the actual file itself
            path = inputPath.rfind('\\')
            #sets the path and drops the file name and extension
            fullPath = inputPath[:path]
            #change directory so the file can be parsed by pandas
            os.chdir(fullPath)
            break
        else:
            print('Invalid file extension. Please provide a valid excel file in .xls or .xlsx')
            continue
    
    #Reads the excel file with pandas
    xl = pd.ExcelFile(inputPath)
    
    #Parse sheet1 from the excel file into a dataframe
    df1 = xl.parse('Sheet1')
    
    #Creates a set of the first and last names from the sheet, which makes them unique.
    names = set((df1['First'] + ' ' + df1['Last']))
       
    #create directories based on unique usernames
    #outputs to log.txt with success or failure messages
    for i in names:
        try: 
            os.mkdir(i)
            print('Director for user', i, 'has been created.')
            logLine = str('\nSuccess! Director for user ' + i + ' has been created.')
            f = open('log.txt', 'a')
            f.write(logLine)
            f.close()
        except Exception:
            print('Directory', i, 'already exists and will not be created.')
            logLine = str('\nFailure! Directory ' + i + ' already exists and will not be created.')
            f = open('log.txt', 'a')
            f.write(logLine)
            f.close()
            pass
    
    #Iterates through all rows moving files into the correct directory based off the name columns from the sheet
    for index, row in df1.iterrows():
        try:
            #Moves files matching the persons first and last name into their newly created directory 
            shutil.move(str(fullPath+'/'+row['document']), str(fullPath+'/'+row['First'] + ' ' + row['Last']+'/'+row['document']))
            print('Moving file', row['document'], 'to location', str(fullPath+'/'+row['First'] + ' ' + row['Last']+'/'+row['document']))
            logLine = str('\nSuccess! Moving file ' + row['document'] + ' to location ' + str(fullPath+'/'+row['First'] + ' ' + row['Last']+'/'+row['document']))
            f = open('log.txt', 'a')
            f.write(logLine)
            f.close()
        except Exception:
            print('Error moving file', row['document'])
            logLine = str('\nFailure! Error moving file ' + row['document'])
            f = open('log.txt', 'a')
            f.write(logLine)
            f.close()
            pass

    #Section of container classes
    #set created above as variable 'names'
    print(names)
    
    #string created above 'inputPath'
    print(inputPath)
    
    #Dictionary class containing the data from the excel file
    nameDict = df1.to_dict()
    
    #Prints the 'First' keys from the dictionary
    print(nameDict['First'])    
      
    #List creation from the dictionary
    for key, value in nameDict.items():
        listExample = [key, value]
    print(listExample)
    
    #Tuple construction from the list just created
    tupExample = [()]
    for d in range(0, 10):
        tupExample.append(listExample[1][d])
    tupExample = tuple(tupExample)
    print(tupExample)
    
main()