# -*- coding: utf-8 -*-
"""
Created on Fri Apr  6 13:17:20 2018

@author: OYB6195
"""
import re
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")#.GetNamespace("MAPI")

inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6) 
                                    
folders = inbox.Folders
for folder in folders:
    if str(folder) =='Уволенные':
        CurFolder = folder.Items

def GetMailData():
    a = []
    for mail in CurFolder:
        file = mail.attachments
        for item in file:
            filePath='d:\\Python\\Mail\\'+str(item)
            item.SaveAsFile(filePath)
            f = open(filePath, 'rb')
            lines = f.read().decode("utf-16")
            users = re.findall(r'[A-Za-z]{3}\d{4}',lines)
            a.append([mail.SentOn,users])
    return a
def SendMail(Mail):
    newMail = outlook.CreateItem(0)
    newMail.Subject = 'Уволенные в QlikView'
    newMail.Body = Mail
    recipient='OYB6195@yum.com'
    newMail.To = recipient
    #newMail.to = 
#    newMail.Display()
    newMail.Send()
    