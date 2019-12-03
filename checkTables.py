#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Wed Nov 27 09:32:25 2019

@author: rguerra
"""

def checkTables( fn ):
    """Loads a workbook, checks for tables on workbook sheets, and saves them to dictionaries keyed by sheetname
    
    Keyword arguments:
        fn -- file name of Workbook file to be opened
    Returned arguments:
        sheetDict -- returns a dictionary keyed by sheet name. In each key are dictionaries keyed by table name; 
                     the table dictionaries contain the location of the table (row:col span) and the contents of the cells
    
    """
    from openpyxl import load_workbook as ldwb
    sheetDict = {}  # initialize dictionary to-be keyed by sheet name
    wb = ldwb( fn ) # load workbook
    for ws in wb:   # for all worksheets in workbook
        if bool(ws._tables):        # if current worksheet has tables
            tableDict = {}  # initialize dictionary to hold data about the tables located in worksheet ; dictionary is keyed by table name
            for tbl in ws._tables:  # for all tables in worksheet
                rowData = []        # re/initialize list for holding contents of table rows
                for row in ws[tbl.ref]: # for all rows in current table
                    cols = []           # re/initialize list for holding contents of table cols
                    for col in row:     # for all cols in row
                        cols.append(col.value) # append to end of list the current cell val
                    rowData.append(cols)       # append to end of row list the value of all cells in row 
                # after saving all values in current table, add entry to table dictionary with all table infos
                tableDict[tbl.name] = { 'location': tbl.ref, 'contents': rowData }
            sheetDict[ws.title] = tableDict    # once all tables are added to table dictionary, place dict into sheet dictionary
        else:
            sheetDict[ws.title] = {} # do not save any info if there are no tables
    return sheetDict 
