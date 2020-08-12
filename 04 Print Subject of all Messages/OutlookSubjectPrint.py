# -*- coding: utf-8 -*-
"""
Created on Mon Aug 10 15:46:00 2020

@author: rahul
"""

import win32com.client

session = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = session.GetDefaultFolder(6) 
messages = inbox.Items

message = messages.GetFirst ()
while message:
  print (message.Subject)
  message = messages.GetNext ()
  
class Folder (object):
  def __init__ (self, folder):
    self._folder = folder
  def __getattr__ (self, attribute):
    return getattr (self._folder, attribute)
  def __iter__ (self):
   
    messages = self._folder.Messages
    message = messages.GetFirst ()
    while message:
      yield message
      message = messages.GetNext ()

if __name__ == '__main__':
  import win32com.client
  session = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
  constants = win32com.client.constants
  session.Logon ()
  
  sent_items = session.GetDefaultFolder(5)
  message = messages.GetFirst ()
  while message:
      print (message.Subject)
      message = messages.GetNext ()