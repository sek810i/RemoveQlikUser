# -*- coding: utf-8 -*-
"""
Created on Wed Apr 11 16:53:21 2018

@author: OYB6195
"""

import ReadMail as RD
import ReadExcel as RE

#RD.SendMail('test')

md = RD.GetMailData()
print('Данные из писем получены')
ed = RE.GetQlikUser()
print('Данные из EXCEL получены')

text=''

for mailUser in md:
    for user in mailUser[1]:
        if user.lower() in ed:
            text+=user.lower()+' из письма от '+ str(mailUser[0])+' \n'
            
if text == '':
    print('Нет пользователей для удаления')
else:
    print('Результат отправляется на почту')           
    RD.SendMail('Нужно удалить:\n'+text)

