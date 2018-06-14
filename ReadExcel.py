# -*- coding: utf-8 -*-
"""
Created on Thu Apr 12 11:39:45 2018

@author: OYB6195
"""

import re
import win32com.client

def FindUser(user, list):
    for item in list:
        if item == user:
            return True
    return False

def GetQlikUser():
    appExcel = win32com.client.Dispatch("Excel.Application")
    
    excel = appExcel.Workbooks.Open('\\\\mosappqv04\\QlikUsers\\QlikUsers.xlsx')
    
    sheet = excel.Worksheets(1)
    dataList = []
    cell = 'null'
    i=1
    while str(cell) != 'None':
        cell = sheet.Cells(i,3)
        #pr = sheet.Cells(i,3)
        
        if str(cell) != 'None' and not FindUser(str(cell),dataList):
            #print(str(cell),str(pr))
            dataList.append(str(cell).replace('TRI-INTL\\','').lower())
        i+=1
    return dataList
