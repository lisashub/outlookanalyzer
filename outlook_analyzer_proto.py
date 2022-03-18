#Prototype Script for Pie-Chart of Top Unread Senders

def unique (list1): #custom function to identify unique senders
    unique_list = []
    for item in list1:
        if item not in unique_list:
            unique_list.append(item)
    return unique_list


import win32com.client #core extraction library
from tqdm import tqdm # library to display extraction progress bar
import pandas as pd # library to tabulate data and generate plot/image
from PIL import Image #library to display image (temporary if we build web frontend)


#Assign Extraction Variables
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
messages = inbox.Items #variable for items in inbox

#Assign Output File Variables
file3 = open("unread_senders.txt", "w+", encoding = "utf-8") #creates data file


#Create a raw list of unread emails consisting of their senders address only
unread_senders = []
print("Retrieving Messages:")
for item in tqdm(messages):
    if (item.Unread == True):
        sender = item.SenderEmailAddress
        unread_senders.append(sender)

#Sends raw list of unread senders to a function that indentifies only the unique sender list
unique_senders = unique(unread_senders)

#Prints some ouputs to Command Line
print("\n")
print("Number of Unread Emails: ", len(unread_senders)) 
print("Number of Unique Unread Mail Senders: " , len(unique_senders))

#Creates sender dictionary consisting of sender addresses and count of unread emails
unique_sender_dict = {}
for sender in unique_senders:
    unique_sender_dict[sender] = unread_senders.count(sender)

#Sorts dictionary according to number of unread emails
unique_sender_list = sorted(unique_sender_dict.items(), key = lambda x:x[1], reverse = True) # sorts dict and stores as list
sorted_unique_sender_dict = dict(unique_sender_list) #converts list back to dict

i = 0 # counter variable to help count collected senders
for item in sorted_unique_sender_dict.items(): #write's top nth senders and email count to file
    if (i<10):
        print(item[0], "\t", item[1], file = file3)
        i = i + 1
file3.close()   
 
sender_table = pd.read_table('unread_senders.txt', sep = '\t', header = None)
print("Pandas Table of Top 10 Senders: ", "\n", sender_table)

#Creates pie plot
plot = sender_table.groupby([0]).sum().plot(kind='pie', y=1, labeldistance=None, autopct='%1.0f%%', title="Senders of Unread Emails")
plot.legend(bbox_to_anchor=(1,1)) #Sets legend details
plot.set_ylabel("Senders") #Set label detail
plot.figure.savefig("sender_plot.jpg", bbox_inches='tight') #saves plot locally
im = Image.open("sender_plot.jpg") #displays plot in default photo viewer; can be moved to web app
im.show()