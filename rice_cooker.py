#!/usr/bin/env python3
#jiraWriter2.py

"""
Description: This script performs data organization of a proprietary Excel (.xlsx) file. Data from the RICEW Tracker and Conversion Plan is used to create a Jira-importable file of requested issues
Author: Jonathan P. Stuecker
Company: PricewaterhouseCoopers Advisory Services LLC
Date: 25 July 2023
License: TBD, Proprietary
Version: 1.0
Dependencies: pandas, csv, random, datetime, numpy, os
Python version: 3.11.4
"""

import pandas as pd
import csv
import random
import datetime
import numpy as np
import os

#global riceTracker, hcmRiceTracker, columns, dates, riceIds
#riceTracker = "Project Orion RICEW Tracker.xlsx"
#hcmRiceTracker = "HCM - Status Summary"
#columns = ['RICE ID', 'Description']
#dates = ['FS Planned Completion Date', 'FS Planned Approval Date', 'TS Planned Completion Date', 'ERP Planned Build Date', 'FUT Planned Completion Date']
#riceIds = ['AM-FF-043', 'AM-FF-044', 'BN-FF-016', 'AM-FF-039', 'PY-FF-051']


def randIssueID():      #Generate random number to be Issue ID
    return random.randrange(100000000000,999999999999)



def generate(excelFile, excelSheet, cols_vals, cols_dates, ids, flag):
#def generate(**kwargs):
#    for key, value in kwargs.items():
#        print("{0} = {1}".format(key, value))

    #--------Subtask Maps----------
    subtask_map_conversions = {'Create/Update Conversion Spec': 'FSD & Mapping Date', \
    'Define Data Validation Workbook Procedures': '', \
    'Build Legacy Conversion Extract Programs': 'Extract', \
    'Execute Legacy Data Preparation/Cleansing': '', \
    'FSD & Mapping Date': '', \
    'Extract Data from Legacy System': 'Extract', \
    'Transform Data into Oracle': 'Transform Data (ERP)', \
    'Profile Extract File': 'Profiling\n(2 Rounds)', \
    'Final Extract File': 'Final Extract', \
    'Final Transform data': 'Final Transform Data (ERP)', \
    'Final Profile': 'Final Profile', \
    'Mapping & Approval': 'Mapping & Approval', \
    'Load Extract File into Oracle': 'Data Load', \
    'Generate Reconciliation Report and Conversion Stats': 'Recon PwC', \
    'Data Reconciliation Sign Off': 'AMC Validation'}

    subtask_map_rief = {"Functional Spec Complete": "FS Planned Completion Date",
                    "Functional Spec Approval": "FS Planned Approval Date",
                    "Technical Spec Complete": "TS Planned Completion Date",
                    "Build Complete": "ERP Planned Build Date",
                    "FUT Complete": "FUT Planned Completion Date"}
    #--------Subtask Maps----------------------------------------------
    #------------------------------------------------------------------
    #--------Check Number of Columns-----------------------------------
    if flag == "RIEF":
        #RIEF
        map = subtask_map_rief
        map = {v: k for k, v in map.items()}
    elif flag == "Conversion":
        #CONVERSIONS
        map = subtask_map_conversions
        map = {v: k for k, v in map.items()}
    else:
        raise Exception("Please make a selection...")
    #--------Check Number of Columns-----------------------------------
    #------------------------------------------------------------------
    ws = pd.read_excel(excelFile, sheet_name=excelSheet)    #Load worksheet

    HEADER = []
    subtask_date_vals = {}  #unused
    issueIDs = {}           #Dictionary of issue IDS with Rice ID as key

    importTasks = []        #Will contain rows for .csv file
    for col in cols_vals:   #Iterate through Column values (not dates) to prepare header
        HEADER.append(col)
    HEADER.extend(["Due Date", "Issue ID", "Parent ID", "Subtask Number", "Subtask"])     #Final values in each row of .csv

    importTasks.append(HEADER)  #Header should be the first row in .csv file

    #---------Begin----------------
    for riceid in ids:      #Iterate through Rice IDs which were input
        i = 0
        parentID = randIssueID() #Generate random issue ID for parent task
        newRow = []         #Clear newRow variable which will be added to importTasks list
        
        #row = ws.loc[ws["RICE ID"]==riceid]     #Locate the row corresponding to the Rice ID in question
        #------------
        #CHANGE SO THAT RICE ID MUST BE THE FIRST VALUE IN SEARCH
        #------------

        row = ws.loc[ws[cols_vals[0]] == riceid]


        #Fetch values from selected columns, insert into newRow list
        for val in cols_vals:   #Iterate through requested values
            try: newRow.append(row[val].values[0])
            except: newRow.append(row[val])

        #Fetch due date of parent task
        finalDate = cols_dates[-1]  #This is the due date of the parent task
        try: finalDate = pd.to_datetime(str(row[finalDate].values[0])).strftime('%Y-%m-%d')  #Fetch value
        except: finalDate = None    #If date does not exist, set to None
        newRow.append(finalDate)   #Add parent task due date
        
        newRow.append(parentID)  #Next value is issueID
        newRow.append("None")   #Parent ID is None for Parent Task
        newRow.append(i)        #Procedural task, add subtask index
        newRow.append("Parent")   #Parent

        importTasks.append(newRow)  #Add parent task to importTasks list

        #Add subtasks
        for subtask in cols_dates:
            i += 1
            #subtask_description = map[subtask]     #CHECK LATER
            newRow = []

            for val in cols_vals:   #Iterate through requested values
                newRow.append(row[val].values[0])
            
            dueDate = pd.to_datetime(str(row[subtask].values[0])).strftime('%Y-%m-%d')  #Fetch value
            newRow.append(dueDate)
            newRow.append(None)
            newRow.append(parentID)
            newRow.append(i)
            newRow.append(subtask)

            importTasks.append(newRow)

    return importTasks





