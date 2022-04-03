#Example Script for identifying flagged emails with a visualization

import win32com.client #core extraction library
from tqdm import tqdm # library to display extraction progress bar
import pandas as pd # library to tabulate data and generate plot/image
import numpy as np
import dataframe_image as dfi


#Assign Extraction Variables
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
messages = inbox.Items #variable for items in inbox

# #Assign Output File Variables
# file1 = open("flagged_email.txt", "w+", encoding = "utf-8") #creates data file


#Create a raw list of flagged emails with associated details

flagged_messages_list = []

i = 0 # counter variable to help count collected senders
print("Retrieving Messages:")
for item in tqdm(messages):

    if (item.FlagRequest == "Follow up"):
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

        # Add each dict to list
        flagged_messages_list.append(flagged_messages_dict)

        # print(i)
        # print(subject)
        # print(sender_email)
        # print(received_time)


#Prints some ouputs to Command Line
print("\n")
print("Number of Flagged emails: " , i)
print(flagged_messages_list)

# Create a dataframe from the list of dictionaries
df = pd.DataFrame(flagged_messages_list)

df_styled = df.style.background_gradient() #adding a gradient based on values in cell
# Export the data frame as an image
dfi.export(df_styled,"flagged_email_list.png")
