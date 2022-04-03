# https://python-forum.io/thread-24810.html

import win32com.client
import win32com
import os
import datetime
 
outlook = win32com.client.Dispatch("outlook.Application").GetNameSpace("MAPI")
inbox=outlook.GetDefaultFolder(6)
messages = inbox.items
 
for message in messages:
    if message.ReceivedTime.date() == datetime.date(2020, 3, 5):
        subject = message.Subject
        if subject == "Test Message":
            print(" ")
            subject=message.Subject
            print ("Subject:", subject) 
            sender = message.Sender
            print ("    Sender  :", sender)
            print(" ")
            timesent = message.senton.strftime("%m/%d/%Y %H:%M:%S")
            print("    Sent    :", timesent)
            receipttime = message.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S")
            print("    Received:",  receipttime)
    message.Close(0)