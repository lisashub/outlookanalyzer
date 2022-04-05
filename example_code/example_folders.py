# https://www.linuxtut.com/en/8a285522b0c118d5f905/

import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

accounts = outlook.Folders

print("root (Number of accounts=%d)" % accounts.Count)
for account in accounts:
    print("└ ",account)
    folders = account.Folders
    for folder in folders:
        print("  └ ",folder)
        mails = folder.Items
        for mail in mails:
            print("-----------------")
            print("subject: " ,mail.subject)
            # print("From: %s [%s]" % (mail.sendername, mail.SenderEmailAddress))
            # print("Received date and time: ", mail.receivedtime)
            # print("Unread: ", mail.Unread)
            # print("Text: ", mail.body)