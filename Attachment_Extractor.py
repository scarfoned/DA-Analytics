#!/usr/bin/env python
from win32com.client import Dispatch
import os


def Extract_DMS_Report(activate_macro):
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

    dms_subject = ["DMS RTU System Health Report ", "DMS RTU System Health Report", "DMS RTU System Health Reports ", "DMS RTU System Health Reports"]
    aml = ["Yes", "yes", "Yea", "yea", "Sure", "sure"]

    box = outlook.GetDefaultFolder(6)
    inbox = box.Folders["DMS Health Report"]
    all_messages = inbox.Items
    first_message = all_messages.GetLast()
    attached = first_message.Attachments.item(1)
    try:
        attached.SaveASFile("C:/Users/C125238/Desktop/DocCentral" + '\\' + "DMS_Summary.xls")
    except ValueError:
        print("Error")
    if activate_macro in aml:
        ActivateMacro()


def Extract_Specific_Message(subject):
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

    dms_subject = ["DMS RTU System Health Report ", "DMS RTU System Health Report", "DMS RTU System Health Reports ", "DMS RTU System Health Reports"]

    inbox = outlook.GetDefaultFolder(6)
    all_messages = inbox.Items

    for message in all_messages:
        print(message)
        if message.Subject == subject:
            attached = message.Attachments
            attach = attached.item(1)
            if str(attach) == attachedfile:
                attach.SaveASFile("C:/Users/C125238/Desktop/DocCentral" + '\\' + 'DMS_Summary.xls')
                print("Mail Successfully Extracted")
            else:
                print("Wrong Filename")


def get_account():
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts = Dispatch("Outlook.Application").Session.Accounts

    for account in accounts:
        global inbox
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        print("****Account Name**********************************")
        print(account.DisplayName)
        print("***************************************************")
        folders = inbox.Folders

        for folder in folders:
            print("****Folder Name**********************************")
            print(folder)
            print("*************************************************")
            a = len(folder.folders)

            if a > 0:
                global z
                z = outlook.Folders(account.DeliveryStore.DisplayName).Folders(folder.name)
                x = z.Folders
                for y in x:
                    print("****Folder Name**********************************")
                    print(y.name)
                    print("*************************************************")


def ActivateMacro():
    filename = "C:/Users/C125238/Desktop/DocCentral/DA Project Tracker.xlsm"
    excel = Dispatch("Excel.Application")
    excel.Workbooks.Open(filename)
    excel.Application.Run("Clear")
    excel.Application.Quit()
    del excel


def main():
    Extract_DMS_Report(input("Would you like to run the macro to replace the data in DA Project Tracker? "))
    #subject = input("Subject of Email: ")
    #Extract_Specific_Message(subject)
    #get_account()


if __name__ == '__main__':
    main()
