#Example Script for identifying flagged emails with a visualization
# There are many ways to export data frame as an image: https://stackoverflow.com/questions/35634238/how-to-save-a-pandas-dataframe-table-as-a-png/
# This might not be the best solution

import win32com.client #core extraction library
from tqdm import tqdm # library to display extraction progress bar
import pandas as pd # library to tabulate data and generate plot/image
# import numpy as np
import dataframe_image as dfi
from PIL import Image #library to display image (temporary if we build web frontend

#Assign Extraction Variables
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
messages = inbox.Items #variable for items in inbox

# #Assign Output File Variables
# file1 = open("flagged_email.txt", "w+", encoding = "utf-8") #creates data file


#Create a raw list of flagged emails with associated details

flagged_messages_list = []

i = 0 # counter variable to help count of flagged messages
print("Retrieving Messages:")
for item in tqdm(messages):

# https://devblogs.microsoft.com/scripting/how-can-i-determine-the-follow-up-status-of-outlook-emails/
# FlagStatus 1 is completed
# FlagStatus 2 is Marked for follow-up

    try:
        if (item.FlagRequest != "" and item.FlagStatus == 2):
            # Using 'item.FlagRequest != ""' is working better than 'item.FlagRequest == "Follow up"'. It caught other flagged emails.

            i += 1 # counter incrementer

            # Create dict object
            flagged_messages_dict = {}
        
            # Assign value and add to dict
            subject = item.subject
            flagged_messages_dict['subject'] = subject

            sender_email = item.SenderEmailAddress
            flagged_messages_dict['sender_email'] = sender_email

            received_time = item.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S")
            flagged_messages_dict['received_time'] = received_time

            flagged_messages_list.append(flagged_messages_dict)

    except Exception as e:

        print("error extracting details for flagged email:" + str(e))
        # Add each dict to list

        # print(i)
        # print(subject)
        # print(sender_email)
        # print(received_time)

number_of_flagged_follow_up_messages = i

#Prints some ouputs to Command Line
print("\n")
print("Number of Flagged emails: " , number_of_flagged_follow_up_messages)
# print(flagged_messages_list)

if (number_of_flagged_follow_up_messages > 0):

    try:

        # Create a dataframe from the list of dictionaries
        df = pd.DataFrame(flagged_messages_list)

    except Exception as e:
        print("error when creating data frame:" + str(e))

    try:

        df_styled = df.style.background_gradient() #adding a gradient based on values in cell
        # Export the data frame as an image
        dfi.export(df_styled,"flagged_email_list.png")
        im = Image.open("flagged_email_list.png") #displays plot in default photo viewer; can be moved to web app
        im.show()

    except Exception as e:
        print("error when exporting data frame:" + str(e))