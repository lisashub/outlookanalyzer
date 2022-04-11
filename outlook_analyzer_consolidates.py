####


import win32com.client #core extraction library
from tqdm import tqdm # library to display extraction progress bar
import pandas as pd # library to tabulate data and generate plot/image
from PIL import Image #library to display image (temporary if we build web frontend)
from tabulate import tabulate
import matplotlib.pyplot as plt
import dataframe_image as dfi
from wordcloud import WordCloud, STOPWORDS #word cloud generation library
import re #regex library used to clean data
import os 
import json

TEMP_DIR = "C:\WINDOWS\Temp"

# class email_parser:
#     TEMP_DIR = "C:\WINDOWS\Temp"
#     print("test")
    # def __init__(self,db):
    #     self.con = sqlite.connect(db)
    
    # def __del__(self):
    #     self.con.close()
    
    # def gettext(self,text,email_id):
    #     word = text[0].split("\r\n")
    #     spliter = re.compile("\\W*")
    #     dataset = []
    #     splitword = []
    #     for w in word:
    #         if len(w) > 0 :
    #             splitword_temp = [str(i).lower() for i in spliter.split(w) if i != '']
    #             splitword.append(splitword_temp)
    #     for line_id in range(0,len(splitword)):
    #         for w in splitword[line_id]:
    #             dataset.append(email_id)
    #             dataset.append(line_id)
    #             dataset.append(w)
    #             self.insertion("word",dataset)
    #             dataset = []
    


def read_email():
    print("test2")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
    inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
    messages = inbox.Items #variable for items in inbox
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    email_temp = []
    email_body = []
    all_messages_list = []

    messages.Sort("[ReceivedTime]",True)
    for item in tqdm(messages):
        messages_dict = {}

        sender_name = item.SenderName
        messages_dict['SenderName'] = sender_name

        sender_email_address = item.SenderEmailAddress
        messages_dict['SenderEmailAddress'] = sender_email_address

        sent_on = item.SentOn
        messages_dict['SentOn'] = sent_on


        try:
            to = item.To
            messages_dict['SentOn'] = to
        except Exception as e:
            print("error extracting details for flagged email:" + str(e))



        cc = item.CC
        messages_dict['CC'] = cc

        bcc = item.BCC
        messages_dict['BCC'] = bcc

        subject = item.Subject
        messages_dict['Subject'] = subject

        subject = item.Subject
        messages_dict['Subject'] = subject

        body = item.Body
        messages_dict['Body'] = body


        # row_count = self.check_email_exists(f)
        # if row_count > 0:
        #     continue
        # email_id = self.insertion("email",email_temp)
        # self.gettext(email_body,email_id)
        # email_temp = []
        # email_body = []

        # print(email_temp)
        # print(email_body)


        # received_time = item.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S")
        # messages_dict['received_time'] = received_time

        # all_messages_list.append(messages_dict)

        # print(all_messages_list)

        json_object = json.dumps(messages_dict, indent = 4)

        with open("C:\WINDOWS\Temp\sample.json", "w+") as outfile:
            outfile.write(json_object)

def main():
    print(0)
    #Placeholder for where user can input whether extraction is necessary
    read_email()



# def main():
#     outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
#     inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
#     messages = inbox.Items #variable for items in inbox
    
#     #Placeholder for where user can input whether extraction is necessary
#     extract_outlook_information(messages)

# def extract_outlook_information(messages): #to modify as new features required
#     #unread senders X
#     #categories X
#     #flags #
#     #word cloud
    
#     #data file creation
#     sender_data_file = open("unread_senders.txt", "w+", encoding = "utf-8")
#     categories_data_file = open("categories.txt", "w+", encoding = "utf-8")
    
    
#     #additional variable creation
#     unread_senders_raw_list = [] #list variable to capture unread email senders with dupes
#     unread_senders_unique_dict = {} #dictionary variable to capture unique unread senders with counts
#     categories_senders_list = [] #dictionary variable to capture category and sender information
#     flagged_messages_list = []
#     flagged_messages_dict = {}
    
#     categories_counter_int = 0
#     senders_counter_int = 0
#     flagged_counter_int = 0
#     message_counter_int = 0
    
#     error_list = []

#     messages.Sort("[ReceivedTime]",True)
#     for item in tqdm(messages):
#         if (item.Unread == True): #checks and stores undread email info
#             sender = item.SenderEmailAddress
#             unread_senders_raw_list.append(sender)
       
#         if item.Categories: #checks and stores info for emails that have been set with categories (by user)
#             categories_senders_list.append([item.SenderEmailAddress,item.Categories])
#             categories_counter_int = categories_counter_int + 1
        
#         try:
#             if (item.FlagRequest != ""): #checks and stores flag info
             
#              # Assign value and add to dict
#              subject = item.subject
#              flagged_messages_dict['subject'] = subject

#              sender_email = item.SenderEmailAddress
#              flagged_messages_dict['sender_email'] = sender_email

#              received_time = item.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S")
#              flagged_messages_dict['received_time'] = received_time

#              flagged_messages_list.append(flagged_messages_dict)
             
#              flagged_counter_int = flagged_counter_int + 1
        
#         except Exception as e:
#             error_list.append("error extracting details for flagged email:" + str(e))
        
#         message_counter_int = message_counter_int + 1
        
#         if message_counter_int == 500:
#             break

        
#     unread_senders_unique_list = unique(unread_senders_raw_list)
    
#     for sender in unread_senders_unique_list:
#         unread_senders_unique_dict[sender] = unread_senders_raw_list.count(sender)
    
#     unread_senders_unique_list = sorted(unread_senders_unique_dict.items(), key = lambda x:x[1], reverse = True) # sorts dict and stores as list
#     unread_senders_unique_sorted_dict = dict(unread_senders_unique_list) #converts list back to dict

#     senders_counter_int = 0 # counter variable to help count collected senders
#     for item in unread_senders_unique_sorted_dict.items(): #write's top nth senders and email count to file
#         if (senders_counter_int<10):
#             print(item[0], "\t", item[1], file = sender_data_file)
#             senders_counter_int = senders_counter_int + 1
    
#     sender_data_file.close()
#     categories_data_file.close()
    
#     generate_unread_senders_viz()
#     generate_categories_viz(categories_counter_int,categories_senders_list)
#     generate_flagged_viz(flagged_counter_int, flagged_messages_list, error_list)
    
#     word_cloud_extract(messages)
#     word_cloud_display()
    
#     if error_list != []:
#         print("\n")
#         print("Some errors occurred during extraction:")
#         for item in error_list:
#             print(item)

# def generate_unread_senders_viz():
#     sender_table = pd.read_table('unread_senders.txt', sep = '\t', header = None)
#     plot = sender_table.groupby([0]).sum().plot(kind='pie', y=1, labeldistance=None, autopct='%1.0f%%', title="Senders of Unread Emails")
#     plot.legend(bbox_to_anchor=(1,1)) #Sets legend details
#     plot.set_ylabel("Senders") #Set label detail
#     plot.figure.savefig("sender_plot.jpg", bbox_inches='tight') #saves plot locally
#     print("\n","Top 10 Senders of Unread Emails: ", "\n", sender_table)
#     im = Image.open("sender_plot.jpg") #displays plot in default photo viewer; can be moved to web app
#     im.show()


# def generate_categories_viz(categories_counter_int,categories_senders_list):
#     #Pandas dataframe for the counted emails that are categorized
#     data = {'Number of email categories': [categories_counter_int]}
#     df = pd.DataFrame(data)

#     #Removing the axis for matplotlib and creating a visual table of the counted emails that are categorize
#     fig, ax = plt.subplots()
#     ax.axis('off')
#     ax.axis('tight')
#     ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
#     fig.tight_layout()
#     plt.savefig("categories.jpg")
#     plt.clf()

#     #prints a tabulate table using the pandas dataframe
#     print("\n")
#     print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex='never'))
#     print(categories_senders_list)

# def generate_flagged_viz(flagged_counter_int, flagged_messages_list, error_list):
#     if (flagged_counter_int  > 0):
#         try:

#             # Create a dataframe from the list of dictionaries
#             df = pd.DataFrame(flagged_messages_list)

#         except Exception as e:
#             error_list.append("error when creating data frame:" + str(e))

#         try:

#             df_styled = df.style.background_gradient() #adding a gradient based on values in cell
#             # Export the data frame as an image
#             dfi.export(df_styled,"flagged_email_list.png")
            
#             print("\n")
#             print("Flagged Emails")
#             print(tabulate(df, headers = 'keys', tablefmt = 'psql'))
#             # im = Image.open("flagged_email_list.png") #displays plot in default photo viewer; can be moved to web app
#             # im.show()

#         except Exception as e:
#             error_list.append("error when exporting data frame:" + str(e))
            
#             #Prints some ouputs to Command Line
#     print("\n")
#     print("Number of Flagged emails: ", flagged_counter_int)
    
# def word_cloud_extract(messages):
#     messages.Sort("[ReceivedTime]",True)
#     wc_file = open("word_cloud_text.txt", "w+", encoding = "utf-8") #creates data file
#     i = 0
#     for item in messages:
#         if(i<50):
#             print(item.Body, file = wc_file)
#             i = i + 1
#         else:
#             word_cloud_content_clean() #text-cleaning function called
#             wc_file.close()
#             return

# #Removes hyperlink information from email body to produce more meaningful clouds
# def word_cloud_content_clean():
#     wc_content= open("word_cloud_text.txt", "r", encoding = "utf-8").read()
#     wc_content_cleaned = open("word_cloud_text_cleaned.txt", "w+", encoding = "utf-8")
    
#     #Sets hyperlink tags as indices markers
#     sub1 = '<'
#     sub2 = '>'
    
#     #Generates list of indices for each tag        
#     indices1 = [m.start() for m in re.finditer(sub1, wc_content)]
#     indices2 = [m.start() for m in re.finditer(sub2, wc_content)]
    
#     #Prints first line of email text up to first hyperlink tag
#     print(wc_content[0:indices1[0]],file = wc_content_cleaned)

#     #Uses iteration through hyperlink tags to extract and print non-hyperlink text to new
#     #file
#     ix = 0 
#     for i in range(len(indices1)-1):
#         print(wc_content[indices2[ix]+1:indices1[ix+1]], file = wc_content_cleaned)
#         ix = ix + 1
        
#     wc_content_cleaned.close()

# #Generates word cloud from cleaned email text
# def word_cloud_generate():
#     wc_content= open("word_cloud_text_cleaned.txt", "r", encoding = "utf-8").read()
#     stop_words = ["said", "email", "s", "will", "u", "re"] + list(STOPWORDS) #customized stopword list
#     wordcloud = WordCloud(stopwords = stop_words).generate(str(wc_content))
#     return(wordcloud)

# #Displays word cloud to user
# def word_cloud_display():
#     plt.imshow(word_cloud_generate())
#     plt.axis('off')
#     plt.savefig('word_cloud.jpg')

# def unique (list1): #custom function to identify unique senders
#     unique_elements_list = []
#     for item in list1:
#         if item not in unique_elements_list:
#             unique_elements_list.append(item)
#     return unique_elements_list

if __name__ == "__main__":
    main()