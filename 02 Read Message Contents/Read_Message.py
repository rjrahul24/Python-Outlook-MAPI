# -*- coding: utf-8 -*-
"""
Created on Mon Aug 10 14:59:19 2020

@author: rahul
"""


import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) 

messages = inbox.Items
message = messages.GetLast()
body_content = message.body
print(body_content)

#For Reference: GetDefaultFolder(6) 6 = Inbox
#Below are the numbers given to specific folders
'''
3  Deleted Items
4  Outbox
5  Sent Items
6  Inbox
9  Calendar
10 Contacts
11 Journal
12 Notes
13 Tasks
14 Drafts
'''

#Instead of Messages.GetLast() Some other functionalities that can be used are
'''
.GetFirst()
.GetLast()
.GetNext()
.GetPrevious()
.Attachments
'''