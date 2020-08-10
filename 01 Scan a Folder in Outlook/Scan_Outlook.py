import win32com.client
import win32com
import os
import sys

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts;

for account in accounts:
    
    folders = outlook.Folders(account.DeliveryStore.DisplayName)
    specific_folder = folders.Folders
    
    for folder in specific_folder:    
        if(folder.name == "DTE-Bot-Folder"):
            messages = folder.Items
            for single in messages:
                if single.Subject == "Poll for DTE":
                    for recp in single.Recipients:
                        print(recp)
                if single.Subject == "Yes: Poll for DTE":
                    print(single.Sender)
    
print("Finished Succesfully")