from datetime import timedelta
from datetime import date
from datetime import datetime
from dateutil.relativedelta import relativedelta
from fpdf import FPDF
from glob import glob
from matplotlib.backends.backend_pdf import PdfPages
from PyPDF2 import PdfFileMerger
from tabulate import tabulate
from tqdm import tqdm
from wordcloud import WordCloud, STOPWORDS

import argparse
import matplotlib.pyplot as plt
import os
import pandas as pd
import re
import shutil
import subprocess
import sys
import time
import traceback
import win32com.client

# sets real time date as a string for PDF file
NOW = datetime.now()
NOW_DATE = NOW.strftime('%m/%d/%y')

#Global script variables created
ERROR_LIST = []
TIME_STR = time.strftime("%Y%m%d-%H%M%S")
TEMP_FOLDER = "outlookanalyzer"
TEMP_DIR = "C:\\WINDOWS\\Temp" + "\\" + TEMP_FOLDER 

# Check whether the specified path exists or not
is_temp_dir_exist = os.path.exists(TEMP_DIR)

if not is_temp_dir_exist:
  
  # Create a new directory because it does not exist 
  os.makedirs(TEMP_DIR)
  print("The new temp directory is created!")

#Global extract text files created
WORD_CLOUD_CLEANED_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud_text_cleaned.txt"
WORD_CLOUD_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud_text.txt"
UNREAD_SENDERS_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "unread_senders.txt"
CATEGORIES_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "categories.txt"
FLAGGED_EMAIL_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "flagged_email.txt"
IMPORTANT_EMAIL_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "important_email.txt"

#Global image files created
WORD_CLOUD_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud.jpg"
SENDER_PLOT_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "sender_plot.jpg"

# Global pdf file (for ordering purposes)
COVER_PAGE_PDF_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "a001.pdf"
COUNTING_PDF_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "b001.pdf"
CATEGORIES_PDF_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "b002.pdf"
FLAGGED_EMAIL_PDF_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "b003.pdf"
IMPORTANT_EMAIL_PDF_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "b004.pdf"
FINAL_REPORT_PDF_FILE_NAME = "C:\\WINDOWS\\Temp\\" + TIME_STR + "_" + "outlook_analyzer_report.pdf" 

IMAGE_FILE_NAME_DICT = {'blue': {"image_path": "black.jpg", "x": "0", "y": "0", "w": "210", "h": "30"},
                        'icon': {"image_path": "icon.png", "x": "0", "y": "0", "w": "35", "h": "30"},
                        'word_cloud': {"image_path": WORD_CLOUD_IMAGE_FILE_NAME, "x": "-35", "y": "100", "w": "275", "h": "250"},
                        'sender_plot': {"image_path": SENDER_PLOT_IMAGE_FILE_NAME, "x": "0", "y": "65", "w": "210", "h": "100"}}

#Function to generate a list of errors that have occurred during program execution; printed at end of run
def append_to_error_list(function_name, error_text, optArg = None): #added optional argument for more detail
    ERROR_LIST.append("function: " + function_name + " | " +  "error: " + error_text)

#Function to extract relevant Outlook information from user's desktop client
def extract_outlook_information(max_email_number_to_extract_input,date_start_input,date_end_input): #to modify as new features required

    #Connection to Outlook object model established
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    todo_folder = outlook.GetDefaultFolder(28) #outlook.GetDefaultFolder(28) is for the todo/flagged items
    messages = inbox.Items
    todo_items = todo_folder.Items

    #Creates data file variables to store extracted information
    sender_data_file = open(UNREAD_SENDERS_DATA_FILE_NAME, "w+", encoding = "utf-8")
    categories_data_file = open(CATEGORIES_DATA_FILE_NAME, "w+", encoding = "utf-8")
    flagged_email_data_file = open(FLAGGED_EMAIL_DATA_FILE_NAME, "w+", encoding = "utf-8")
    important_email_data_file = open(IMPORTANT_EMAIL_DATA_FILE_NAME, "w+", encoding = "utf-8")

    #Creates intermediate list and dictionary structure variable to store extracted information
    category_list = []
    counting_dict = {}
    flagged_messages_list = []
    important_messages_list = []
    unread_senders_raw_list = []
    unread_senders_unique_dict = {}
   
    #Creates variables to count key inbox properties
    categories_counter_int = 0
    flagged_counter_int = 0
    important_count_int = 0
    message_counter_int = 0
    message_read_counter_int = 0
    message_unread_counter_int = 0
    number_of_times_categories_assigned_counter_int = 0

    #Establishes how many months or days back the script should look for emails
    if date_end_input[-1] == "m": #Value will be "m" if user enters range in months
        month_int = int(date_end_input[0:-1])

        if month_int == 0: #Value will be 0 if user does not enter a date rage; default used
            end_date = datetime.now()
        else:
           end_date = datetime.now() - relativedelta(months=+month_int)

    elif date_end_input[-1] == "d": #Value will be "d" if user enters range in days
        day_int = int(date_end_input[0:-1])
 
        if day_int == 0:
            end_date = datetime.now()
        else:
            end_date = datetime.now() - timedelta(days=day_int)

    #Establishes how many months or days from the present the script should look for emails; see comments above for value descriptions
    if date_start_input[-1] == "m":
        month_int = int(date_start_input[0:-1])

        if month_int == 0:
            start_date = date.today()
        else:
           start_date = date.today() - relativedelta(months=+month_int)

    elif date_start_input[-1] == "d":
        day_int = int(date_start_input[0:-1])
 
        if day_int == 0:
            start_date = date.today()
        else:
            start_date = date.today() - timedelta(days=day_int)

    #Converts email extraction date inputs into string format for filtering messages
    date_start_str = start_date.strftime('%m/%d/%Y %H:%M %p')
    date_end_str = end_date.strftime('%m/%d/%Y %H:%M %p')

    #Combines end and start date inputs into a range
    filtered_messages = messages.Restrict("[ReceivedTime] >= '" + date_start_str + "' AND [ReceivedTime] <= '" + date_end_str + "'")

    print("Extracting email messages:")
    
    #Iterates through inbox items and extracts relevant information
    filtered_messages.Sort("[ReceivedTime]",True)
    for inbox_item in tqdm(filtered_messages): # Displays tdqm progress bar during iteration

        #Unread email metric logic
        try:
            
            if (inbox_item.UnRead == True):

                message_unread_counter_int = message_unread_counter_int + 1
    
                sender = return_sender(inbox_item)

                unread_senders_raw_list.append(sender)

            else:
                message_read_counter_int = message_read_counter_int + 1
                
        except AttributeError as e:
            
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
            
            path = os.environ['USERPROFILE']+"\AppData\Local\Temp\gen_py"
            
            if os.path.isfile(path):
                
                shutil.rmtree(path)               
            
            else:
                raise Exception
                
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
        
        #Assigned category metric logic 
        try:
            if inbox_item.Categories:
                item_categories = inbox_item.Categories.split(",") #Splits multiple categories if applicable

                categories_counter_int = categories_counter_int + 1
                for category in item_categories:
                    category_list.append(category.strip())
                    number_of_times_categories_assigned_counter_int = number_of_times_categories_assigned_counter_int + 1
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

        #High importance email flag metric logic
        if (inbox_item.Importance == 2 and inbox_item.Class == 43): #Importace is 2 if item is marked as High Importance
            important_messages_dict = {} # To organize item information
            try:
                subject = inbox_item.Subject
                clean_subject = cleanup(subject) #Removes invisible white space/pointers that pandas cannot handle
                subject = [clean_subject.encode("utf-8").strip()] #Encodes subject as utf8 to handle special characters

                important_messages_dict['subject'] = subject
            except Exception as e:
                append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
        
            try:

                email_class = inbox_item.Class
                important_messages_dict['Class'] = email_class
            except Exception as e:
                append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

            try:
                received_time = inbox_item.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S")
                important_messages_dict['ReceivedTime'] = received_time
            except Exception as e:
                append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

            #Retrieves more readable sender data if sender is within internal MS exchange
            if inbox_item.SenderEmailType == "EX":
                try:
                    sender_email = inbox_item.Sender.GetExchangeUser().PrimarySmtpAddress
                except Exception as e:
                    append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
            else:
                try:
                    sender_email = inbox_item.SenderEmailAddress
                except Exception as e:
                    append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

            important_messages_dict['SenderEmailAddress'] = sender_email

            #Adds the attributes for each item to the important_messages_dict
            important_messages_list.append(important_messages_dict)

            important_count_int = important_count_int + 1

        # Count number of messages in loop
        message_counter_int = message_counter_int + 1
        #End of inbox item iteration loop
        
        #Checks if max number of emails has been reached       
        if message_counter_int >= int(max_email_number_to_extract_input):
            break

    #Iterates through to-do items and flagged email; see comments associated with similar code above for additional insight 
    tasks = todo_items.Restrict("[Complete] = FALSE")
    
    print("Extracting tasks/flagged items:")
    for task in tqdm(tasks):      
        flagged_messages_dict = {} 

        #Captures info for all tasks
        try:
            subject = task.Subject
            clean_subject = cleanup(subject) #Removes invisible white space / pointers that pandas cannot handle
            subject = [clean_subject.encode("utf-8").strip()] # Common for subjects to have emoji or utf8 data so need encode as utf8
            flagged_messages_dict['subject'] = subject
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

        try:
            item_class = task.Class
            flagged_messages_dict['Class'] = item_class
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

        #Captures email-related task info
        if task.Class == 43:

            try:
                received_time = task.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S")
                flagged_messages_dict['ReceivedTime'] = received_time
            except Exception as e:
                append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

            sender_email = return_sender(task)

            flagged_messages_dict['SenderEmailAddress'] = sender_email

        # Add the attributes for each task to the flagged_messages_dict
        flagged_messages_list.append(flagged_messages_dict)

        # Keep track of the numbers of tasks/flagged items
        flagged_counter_int = flagged_counter_int + 1

    unique_category_count_int = len(unique(category_list))

    # Add the various count metrics to counting_dict
    counting_dict['total messages'] = message_counter_int
    counting_dict['total unread message'] = message_unread_counter_int
    counting_dict['total read messages'] = message_read_counter_int
    counting_dict['total flagged messages'] = flagged_counter_int
    counting_dict['total categorized messages'] = categories_counter_int
    counting_dict['total times categories used'] = number_of_times_categories_assigned_counter_int
    counting_dict['total unique categories'] = unique_category_count_int
    counting_dict['total email marked as important'] = important_count_int

    if message_counter_int < 1 or message_unread_counter_int < 1:
        print("Insufficient data to analyze unread email.")
    else:
        unread_senders_data_gen(unread_senders_raw_list, unread_senders_unique_dict,sender_data_file)
        generate_unread_senders_viz()

    if message_counter_int < 1 or categories_counter_int < 1:
        print("Insufficient data to analyze categories.")
    else:
        category_data_gen(category_list, categories_data_file)
        
        # Categories pdf
        title_str = "Categories"
        figure_column_list = ['Category Name', 'Count']
        convert_csv_to_df_to_figure_to_pdf(CATEGORIES_DATA_FILE_NAME,title_str,figure_column_list,CATEGORIES_PDF_FILE_NAME)
    
    if flagged_counter_int < 1:
        print("Insufficient data to analyze flagged messages.")
    else:
        # Converts collected data into a text file for flagged messages 
        build_text_with_subject_senderemail_receivedtime(flagged_messages_list, flagged_email_data_file)

        # Flagged email and todo items
        title_str = "Flagged email / Todo"
        figure_column_list = ["Subject","Sender Email","Date"]
        convert_csv_to_df_to_figure_to_pdf(FLAGGED_EMAIL_DATA_FILE_NAME,title_str,figure_column_list,FLAGGED_EMAIL_PDF_FILE_NAME)

    if message_counter_int < 1 or important_count_int < 1:
        print("Insufficient data to analyze important messages.")
    else:
        # Converts collected data into a text file for important messages 
        build_text_with_subject_senderemail_receivedtime(important_messages_list, important_email_data_file)

        # Email that came in as "Important"
        title_str = "Email sent as Important"
        figure_column_list = ["Subject","Sender Email","Date"]
        convert_csv_to_df_to_figure_to_pdf(IMPORTANT_EMAIL_DATA_FILE_NAME,title_str,figure_column_list,IMPORTANT_EMAIL_PDF_FILE_NAME)

    # Count summary
    title_str = "Counts (" + "Start: " + date_start_str + " | End: " + date_end_str + ")"
    figure_column_list = ['Item', 'Count']
    convert_dict_to_df_to_figure_to_pdf(counting_dict, title_str,figure_column_list,COUNTING_PDF_FILE_NAME)

    if message_counter_int < 1 or message_unread_counter_int < 1:
        print("Insufficient data to analyze message content for word cloud.")
    else:
        word_cloud_extract(messages)
        generate_word_cloud_viz()

    create_pdf_cover_page(message_counter_int,message_unread_counter_int)

def return_sender(outlook_object):
    """ Returns sender email address """

    if outlook_object.Class == 43: #Class "43" is assigned to VBA MailItem objects (i.e. regular emails):https://docs.microsoft.com/en-us/office/vba/api/outlook.olobjectclass
        if outlook_object.SenderEmailType == "EX": # SenderEmailType "EX" is assigned to MailItems received from internal MS Exchange
            sender = outlook_object.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            sender = outlook_object.SenderEmailAddress
    else:
        sender = outlook_object.SenderEmailAddress

    return sender

#Extracts word cloud information from most recent 50 messages
def word_cloud_extract(messages):
    try:
     wc_file = open(WORD_CLOUD_FILE_NAME, "w+", encoding = "utf-8") #creates data file
     i = 0
     while (i<50):
         print(messages[i].Body, file = wc_file)
         i = i + 1
     word_cloud_content_clean() #text-cleaning function called
     wc_file.close()
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

#Generates data for unread senders plot
def unread_senders_data_gen(unread_senders_raw_list,unread_senders_unique_dict,sender_data_file):
    try:
        unread_senders_unique_list = unique(unread_senders_raw_list)
        
        for sender in unread_senders_unique_list:
            unread_senders_unique_dict[sender] = unread_senders_raw_list.count(sender)
            
        unread_senders_unique_list = sorted(unread_senders_unique_dict.items(), key = lambda x:x[1], reverse = True) #Sorts dict and stores as list
        unread_senders_unique_sorted_dict = dict(unread_senders_unique_list) #Converts list back to dict
        senders_counter_int = 0
        
        #Collects data for top 10 senders of unread email
        for item in unread_senders_unique_sorted_dict.items():
            if (senders_counter_int<10):
                print(item[0], "\t", item[1], file = sender_data_file)
                senders_counter_int = senders_counter_int + 1
        sender_data_file.close()
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
    
#Generates unread senders table and image file
def generate_unread_senders_viz():
    try:
        sender_table = pd.read_table(UNREAD_SENDERS_DATA_FILE_NAME, sep = '\t', header = None)
        plot = sender_table.groupby([0]).sum().plot(kind='pie', y=1, labeldistance=None, autopct='%1.0f%%', title="Senders of Unread Emails")
        plot.legend(bbox_to_anchor=(1,1)) #Anchors plot legend to right of plot
        plot.set_ylabel("Senders")
        plot.figure.savefig(SENDER_PLOT_IMAGE_FILE_NAME, bbox_inches='tight') #Saves plot locally
        print("\n","Top 10 Senders of Unread Emails: ", "\n", sender_table)
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

# Currently used by Flagged email and Important email metric
def build_text_with_subject_senderemail_receivedtime(messages_list, email_data_file):
    """ Reformats item text data to UTF-8 """
    # Have to generate a text file to decode the utf8 data
    try:

        for item in messages_list:

            if item['Class'] == 43: 
                print('\n'.join(s.decode('utf-8', 'ignore') for s in item['subject']),"\t",item['SenderEmailAddress'], "\t", item['ReceivedTime'], file = email_data_file)
            else: # items from the todo folder do not have sender email address or received time
                print('\n'.join(s.decode('utf-8', 'ignore') for s in item['subject']),"\t", "-", "\t", "-", file = email_data_file)

        email_data_file.close()
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

#Generates categories metric data
def category_data_gen(category_list,categories_data_file):

    """ Generates category data by using the 'category_list', runs it through the 'unique' function,
    saves results into a dict, and then prints it on a text file with the variable 'categories_data_file' """

    category_dict = {} #dictionary variable to capture email category
    
    try:
        unique_categories = unique(category_list) #sends category list to function named "unique" and saves list of unique values to variable

        for category in unique_categories: #loops through unique categories and counts occurrences; saves results into category_dict
            category_dict[category] = category_list.count(category)
        
        categories_counter_int = 0
        for item in category_dict.items():
            print(item[0], "\t", item[1], file = categories_data_file)
            categories_counter_int = categories_counter_int + 1
            
        print('\n')
        print('Number of categories:', len(unique_categories))  # print total number or emails categorize
            
        categories_data_file.close()
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

def convert_dict_to_df_to_figure_to_pdf(metric_dict, title_str, columns_list, pdf_file_name):
    """ Converts a dict into a data frame and then to a figure and exports it as a single pdf """
    try:
        #Pandas dataframe for passed in dict
        df = pd.DataFrame(metric_dict.items(), columns=columns_list)   

        fig, ax =plt.subplots()
        plt.title(title_str, backgroundcolor='black', color='white')
        ax.axis('tight')
        ax.axis('off')
        table = ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
        table.scale(1, 2)

        pp = PdfPages(pdf_file_name)
        pp.savefig(fig, bbox_inches='tight')

        pp.close()
        
        #prints a tabulate table using the pandas dataframe
        print("\n")
        print(title_str)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex='never'))
        
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

# Code borrowed from https://stackoverflow.com/questions/3444645/merge-pdf-files
def pdf_merge(open_file,output_file_name):
    """ Merges all the pdf files in current directory """
    merger = PdfFileMerger()
    location_to_check_str = TEMP_DIR + "\\" + "*.pdf"
    allpdfs = [a for a in glob(location_to_check_str)]
    [merger.append(pdf) for pdf in allpdfs]
    with open(output_file_name, "wb") as new_file:
        merger.write(new_file)

    merger.close()

    if open_file:
        subprocess.Popen([output_file_name],shell=True)

def create_pdf_cover_page(message_counter_int,message_unread_counter_int):

    """ Add images into a PDF file """
    try:
        # Code for creating first page of PDF report
        pdf = FPDF()
        pdf.add_page()  # adds pdf page
        pdf.set_font('Arial', 'B', 18)  # sets pdf fonts
        pdf.cell(0, 60, 'Outlook Analyzer Report', 0, 0, align='C')  # Puts in title
        pdf.cell(-190, 75, NOW_DATE, 0, 0, align='C') # Puts in real time date


        # Traverse through nested dictionary
        for image_id, image_info in IMAGE_FILE_NAME_DICT.items():
            # for troubleshooting
            # print("\nItem:", image_id)
            
            for key in image_info:
                # for troubleshooting
                # print(key + ':', image_info[key])

                is_image_file_exists = os.path.exists(image_info['image_path'])

                if is_image_file_exists:
                    pdf.image(image_info['image_path'], x=int(image_info['x']), y=int(image_info['y']), w=int(image_info['w']), h=int(image_info['h']))  

        pdf.output(COVER_PAGE_PDF_FILE_NAME, 'F')  # saves pdf
        pdf.open()


    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

def convert_csv_to_df_to_figure_to_pdf(email_data_file,title_str,columns_list,pdf_file_name):
    """ Converts a csv into a data frame and then to a figure and exports it as a single pdf """
    try:
        df = pd.read_csv(email_data_file, sep = "\t", encoding ='utf-8', names=columns_list)
        
        if "Subject" in df.columns:
            for i in range(df.shape[0]): #iterates over rows
                if len(df.at[i, "Subject"]) > 50: #checks if Subject is over 50 chars; seems like happy medium
                    df.at[i, "Subject"] = df.at[i, "Subject"][0:46] + "..." #truncates subject

        fig, ax =plt.subplots(figsize=(12,4))
        plt.title(title_str, backgroundcolor='black', color='white')
        ax.axis('tight')
        ax.axis('off')
        table = ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
        table.scale(1, 2)
        table.auto_set_font_size(False)  #stops automatic text shrink which defaults to True
        table.set_fontsize(8)

        pp = PdfPages(pdf_file_name)
        pp.savefig(fig, bbox_inches='tight')

        pp.close()
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
        
    try:
        print("\n")
        print(tabulate(df, headers = 'keys', tablefmt = 'fancy_grid',showindex='never'))

    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

#Generates word cloud visualization
def generate_word_cloud_viz():
    try:
        wc_cleaned_content_file = open(WORD_CLOUD_CLEANED_FILE_NAME, "r", encoding = "utf-8").read()
        
        #Sets stopwords  for cloud
        stop_words = ["said", "email", "s", "will", "u", "re", "3A", "2F", "safelinks", "reserved", "https"] + list(STOPWORDS) #customized stopword list
        
        #Generates word cloud
        word_cloud = WordCloud(width=800, height=400, stopwords = stop_words).generate(str(wc_cleaned_content_file))
        plt.clf()
        plt.rcParams["figure.figsize"] = (10,8)
        plt.imshow(word_cloud)
        plt.axis('off')
        plt.savefig(WORD_CLOUD_IMAGE_FILE_NAME)
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

#Code borrowed from https://stackoverflow.com/questions/49267999/remove-u202a-from-python-string
def cleanup(inp):
    """ Removes converse characters ("\u202a") and pop directional formatting characters from strings ("\u202c") """
    new_char = ""
    for char in inp:
        if char not in ["\u202a", "\u202c"]:
            new_char += char
    return new_char

def delete_temp_files(type_list):
    """ Goes through directory and removes/cleans the files with specified extensions (i.e .txt, .tmp, .png, etc.) """
    while True:

        for item_type in type_list: # looping over the file type to remove
            dir_list = os.listdir(TEMP_DIR) #returns the list of all files and directories in the specified path

            for item in dir_list: # looping files to remove

                try:
                    if item.endswith(item_type):
                        os.remove(os.path.join(TEMP_DIR, item))
                except Exception as e:
                    append_to_error_list(str(sys._getframe().f_code.co_name),str(e), traceback.format_exc())

        break;

#Removes hyperlink information from email body text to make more meaningful clouds
def word_cloud_content_clean():
    try:
        #Opens extracted email body text and cleansed text storage file
        wc_content= open(WORD_CLOUD_FILE_NAME, "r", encoding = "utf-8").read()
        wc_content_cleaned = open(WORD_CLOUD_CLEANED_FILE_NAME, "w+", encoding = "utf-8")
        
        #Create indices counter variable
        indices_counter_int = 0
        
        #Assigns variables to tags preceding link information
        sub1 = '<http'
        sub2 = '<mail'
        
        #Creates list of indices where opening tags occur in email text body
        indices1_link = [m.start() for m in re.finditer(sub1, wc_content)]
        indices2_mail = [m.start() for m in re.finditer(sub2, wc_content)]
        
        #Combines indices lists
        indices1_link.extend(indices2_mail)
        indices1_link.sort()
        
        #Creates variable to capture indices of closing link tags
        indices2 = []
        
        #Identifies and stores indices values
        for indices in indices1_link:
                end_indices = wc_content[indices:].find('>')
                indices2.append(end_indices+indices)
        
        #Iterates through email body text using identified indices; print non-link data into cleansed file
        while indices_counter_int < len(indices2)-1:
            print(wc_content[indices2[indices_counter_int]+1:indices1_link[indices_counter_int+1]], file = wc_content_cleaned)
            indices_counter_int = indices_counter_int + 1
            
        wc_content_cleaned.close()
    
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

def unique (list1):
    """ Identifies unique elements within a list """
    unique_elements_list = []
    for item in list1:
        if item not in unique_elements_list:
            unique_elements_list.append(item)
    return unique_elements_list

# Code borrowed from https://note.nkmk.me/en/python-check-int-float/
def is_integer_num(n):
    if isinstance(n, int):
        return True
    if isinstance(n, float):
        return n.is_integer()
    return False

def main(argv):  
    """ Runs through the specified inputs that users will enter to get their analysis data for Outlook """
    max_email_number_to_extract_input = 500
    date_start_input = "12m"
    date_end_input = "0m"
    suppress_prompt = False
    open_file = True

    # Initialize parser
    parser = argparse.ArgumentParser()

    # Adding arguments
    parser.add_argument("-n", "--number", help = "Max number of email messages you would like to extract (between 50 and 100000)")
    parser.add_argument("-s", "--start", help = "From how far back would you like to collect and analyze emails in months or days (e.g. 10m, 12d)")
    parser.add_argument("-e", "--end", help = "What's the cutoff for the most recent emails you'd like to collect and analyze in months or days (e.g. 1m, 10d)")
    parser.add_argument("-o", "--output", help = "Location to save the report (must used .pdf extension)")
    parser.add_argument("-O", "--open", help = "Open the report at the end of script running?")


    # Read arguments from command line
    args = parser.parse_args()
        
    if args.number:
        max_email_number_to_extract_input = args.number
        print("number: " +  str(max_email_number_to_extract_input))

        if int(max_email_number_to_extract_input) < 50 or int(max_email_number_to_extract_input) > 100000:
            print("Error: Problem with email number argument.")
            print("Please enter a valid integer between 50 and 100000.")
            exit()

    if args.start:
        date_start_input = args.start
        print("start: " + str(date_start_input))

    if args.end:
        date_end_input = args.end
        print("end: " + str(date_end_input))

    if args.output:
        output_file_name = args.output
        if not output_file_name.lower().endswith('.pdf'):
            print("Error: Problem with output file extension.")
            print("Please ensure that the extension is .pdf")
            exit()
        print("output: " + str(output_file_name))
    else:
         output_file_name = FINAL_REPORT_PDF_FILE_NAME

    if args.open:
        open_file = eval(args.open.capitalize())
        print("open_file: " + str(open_file))

    if args.start or args.end:
        # Check if command line argument are an integer
        if not is_integer_num(int(date_start_input[0:-1])) or not is_integer_num(int(date_end_input[0:-1])):
            print("Error: Problem with date format.")
            print("Please ensure that date is an integer.")
            exit()

        # Check if command line arguments use m or d after integer value
        if not (date_start_input[-1] == "m" or date_start_input[-1] == "d") or not (date_end_input[-1] == "m" or date_end_input[-1] == "d"):
            print("Error: Problem with date format.")
            print("This is not a valid format. Please enter as '##m' or '##d' where d is for days and m is for months  (e.g. 10d or 1m)")
            exit()
  
    # If all three command line arguments are supplied, suppress asking for more input
    if args.number and args.start and args.end:
        suppress_prompt = True

    print("\n")
    print("Welcome to Outlook Analyzer!")

    # If arguments not provided at command line, ask for them
    if not suppress_prompt:
        
        if not args.number:
            #Receives and checks max number of emails to extract
            while True:
                max_email_number_to_extract_input = input("Max number of email messages you would like to extract (between 50 and 100000)? (Hit Enter for default: 500)") or 500

                try:
                    int(max_email_number_to_extract_input)
                    if int(max_email_number_to_extract_input) >= 50 and int(max_email_number_to_extract_input) <= 100000:
                        break;
                except ValueError:
                    print("Please enter a valid integer between 50 and 100000.")

        if not args.start:
            #Receives and checks user input for oldest email cut-off date
            while True:
                date_start_input = input("From how far back would you like to collect and analyze emails in months or days (e.g. 10m, 12d)? (Hit enter for default: 12 months ago)") or "12m"

                try:
                    int(date_start_input[0:-1])
                    if date_start_input[-1] == "m" or date_start_input[-1] == "d":
                        break;
                except ValueError:
                    print("This is not a valid format. Please enter as '##m' or '##d' where d is for days and m is for months  (e.g. 10d or 1m")

        if not args.end:
            #Receives and checks user input for email recency cut-off date
            while True:
                date_end_input = input("What's the cutoff for the most recent emails you'd like to collect and analyze in months or days (e.g. 1m, 10d)? (Hit enter for default: today)") or "0m"

                try:
                    int(date_end_input[0:-1])
                    if date_end_input[-1] == "m" or date_end_input[-1] == "d":
                        break;
                except ValueError:
                    print("This is not a valid format. Please enter as '##m' or '##d' where d is for days and m is for months  (e.g. 10d or 1m")

    # Clean left over temp .pdf files from previous run if exist
    delete_temp_files([".pdf"])

    # Begin email extraction
    extract_outlook_information(max_email_number_to_extract_input,date_start_input,date_end_input)

    pdf_merge(open_file,output_file_name)
    # Clean left over temp .txt, .jpg, .png files from previous run if exist
    delete_temp_files([".txt",".jpg",".png"])

    #Displays any errors generated throughout program execution
    if ERROR_LIST != []:
        print("\n")
        print("Some errors occurred during execution:")
        for item in ERROR_LIST:
            print(item)

if __name__ == "__main__":
    main(sys.argv[1:])
