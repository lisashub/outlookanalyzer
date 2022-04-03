# Example code from: https://www.codeforests.com/2020/06/04/python-to-read-email-from-outlook/
# To retrieve all Outlook items from the folder that meets the predefined condition, you need to sort the items in ascending order:
# https://stackoverflow.com/questions/62588493/python-outlook-extract-mails-by-date-range-gives-a-fixed-lower-limit-irrespecti

import win32com.client
#other libraries to be used in this script
import os
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

for account in mapi.Accounts:
    print(account.DeliveryStore.DisplayName)

inbox = mapi.GetDefaultFolder(6)

# inbox = mapi.GetDefaultFolder(6).Folders["your_sub_folder"]

messages = inbox.Items

# Examples of restricting which email to collect
# received_dt = datetime.now() - timedelta(days=3)
# print (received_dt)
# received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
# messages = messages.Restrict("[ReceivedTime] >= '" + str(received_dt) + "'")
# messages = messages.Restrict("[SenderEmailAddress] = 'contact@codeforests.com'")
# messages = messages.Restrict("[Subject] = 'Sample Report'")

#Let's assume we want to save the email attachment to the below directory
outputDir = r"C:\test"
try:
    for message in list(messages):
        try:

            s = message.sender
            e = message.SenderEmailAddress
            print(s)
            print(e)
            for attachment in message.Attachments:
                attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
                print(f"attachment {attachment.FileName} from {s} saved")
        except Exception as e:
           print("error when saving the attachment:" + str(e))
except Exception as e:
        print("error when processing emails messages:" + str(e))