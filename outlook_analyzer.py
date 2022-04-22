from datetime import timedelta
from datetime import date
from datetime import datetime
from dateutil.relativedelta import relativedelta
from tabulate import tabulate
from tqdm import tqdm
from wordcloud import WordCloud, STOPWORDS


import matplotlib.pyplot as plt
import os
import pandas as pd
import re
import shutil
import sys
import time
import win32com.client




#Global script variables created
ERROR_LIST = []
TEMP_DIR = "C:\WINDOWS\Temp"
TIME_STR = time.strftime("%Y%m%d-%H%M%S")
WORD_CLOUD_CLEANED_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud_text_cleaned.txt"
WORD_CLOUD_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud_text.txt"
WORD_CLOUD_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud.jpg"

#Global extract text files created
UNREAD_SENDERS_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "unread_senders.txt"
CATEGORIES_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "categories.txt"
FLAGGED_EMAIL_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "flagged_email.txt"
IMPORTANT_EMAIL_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "important_email.txt"


#Global image files created
UNREAD_SENDERS_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "sender_plot.jpg"
CATEGORIES_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "categories.jpg"
FLAGGED_EMAIL_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "flagged_email_list.png"
IMPORTANT_EMAIL_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "important_email_list.png"
FLAGGED_EMAIL_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "flagged_email_list.png"
IMPORTANT_EMAIL_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "important_email_list.png"
SENDER_PLOT_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "sender_plot.jpg"
CATEGORIES_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "categories.jpg"
COUNTING_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "counting.jpg"

#Function to generate a list of errors that have occurred during progam execution; printed at end of run
def append_to_error_list(function_name, error_text):
    ERROR_LIST.append("function: " + function_name + " | " +  "error: " + error_text)

#Function to extract relavent Outlook information from user's desktop client
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
    messages.Sort("[ReceivedTime]",True)
    for inbox_item in tqdm(filtered_messages): # Displays tdqm progress bar during iteration
        
        #Unread email metric logic
        try:
            if (inbox_item.UnRead == True):

                message_unread_counter_int = message_unread_counter_int + 1

                if inbox_item.Class == 43: #Class "43" is assigned to VBA MailItem objects (i.e. regular emails):https://docs.microsoft.com/en-us/office/vba/api/outlook.olobjectclass
                    if inbox_item.SenderEmailType == "EX": # SenderEmailType "EX" is assigned to MailItems received from internal MS Exchange
                        sender = inbox_item.Sender.GetExchangeUser().PrimarySmtpAddress
                    else:
                        sender = inbox_item.SenderEmailAddress
                else:
                    sender = inbox_item.SenderEmailAddress
    
                unread_senders_raw_list.append(sender)

            else:
                message_read_counter_int = message_read_counter_int + 1
                
        except AttributeError as e:
            
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
            
            path = os.environ['USERPROFILE']+"\AppData\Local\Temp\gen_py"
            
            if os.path.isfile(path):
                
                shutil.rmtree(path)
                
                message_unread_counter_int = message_unread_counter_int + 1

                if inbox_item.Class == 43: #Class "43" is assigned to VBA MailItem objects (i.e. regular emails):https://docs.microsoft.com/en-us/office/vba/api/outlook.olobjectclass
                    if inbox_item.SenderEmailType == "EX": # SenderEmailType "EX" is assigned to MailItems received from internal MS Exchange
                        sender = inbox_item.Sender.GetExchangeUser().PrimarySmtpAddress
                    else:
                        sender = inbox_item.SenderEmailAddress
                else:
                    sender = inbox_item.SenderEmailAddress
    
                unread_senders_raw_list.append(sender)
                
                continue
            
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

        #Follow-up and high importance email flag metric logic
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
        
        message_counter_int = message_counter_int + 1
        #End of inbox item iteration loop
        
        #Checks if max number of emails has been reached       
        if message_counter_int >= int(max_email_number_to_extract_input):
            break
    
    #Iterates through to-do items; see comments associated with similar code above for additional insight 
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

            if task.SenderEmailType == "EX":
                try:
                    sender_email = task.Sender.GetExchangeUser().PrimarySmtpAddress
                except Exception as e:
                    append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
            else:
                try:
                    sender_email = task.SenderEmailAddress
                except Exception as e:
                    append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

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
    counting_dict['total email maked as important'] = important_count_int

    generate_count_viz(counting_dict, date_start_str, date_end_str)
    unread_senders_data_gen(unread_senders_raw_list, unread_senders_unique_dict,sender_data_file)
    generate_unread_senders_viz()
    category_data_gen(category_list, categories_data_file)
    generate_categories_viz()

    generic_email_data_gen(flagged_messages_list, flagged_email_data_file)
    generic_email_data_gen(important_messages_list, important_email_data_file)

    generate_generic_viz(flagged_counter_int,FLAGGED_EMAIL_DATA_FILE_NAME,FLAGGED_EMAIL_IMAGE_FILE_NAME,"Flagged email / Todo" )
    generate_generic_viz(important_count_int, IMPORTANT_EMAIL_DATA_FILE_NAME,IMPORTANT_EMAIL_IMAGE_FILE_NAME,"Email sent with Important")

    word_cloud_extract(messages)
    generate_word_cloud_viz()
    
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
        plot.figure.savefig(UNREAD_SENDERS_IMAGE_FILE_NAME, bbox_inches='tight') #Saves plot locally
        print("\n","Top 10 Senders of Unread Emails: ", "\n", sender_table)
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))


#Reformats item text data to UTF-8
def generic_email_data_gen(messages_list, email_data_file):

    # Have to generate a text file to decode the utf8 data
    try:
        print("Subject","\t",'SenderEmailAddress',"\t", 'ReceivedTime', file = email_data_file)

        for item in messages_list:

            if item['Class'] == 43: 
                print('\n'.join(s.decode('utf-8', 'ignore') for s in item['subject']),"\t",item['SenderEmailAddress'], "\t", item['ReceivedTime'], file = email_data_file)
            else:
                print('\n'.join(s.decode('utf-8', 'ignore') for s in item['subject']),"\t", "-", "\t", "-", file = email_data_file)

        email_data_file.close()
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))


#Generates categories metric data
def category_data_gen(category_list,categories_data_file):

    category_dict = {} #dictionary variable to capture email category
    
    try:
        unique_categories = unique(category_list) #sends category list to function named "unique" and saves list of unique values to variable

        for category in unique_categories: #loops through unique categories and counts occurrances; saves results into category_dict
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

#Generates counting visualizations
def generate_count_viz(counting_dict, date_start_str, date_end_str):
    try:
        #Pandas dataframe for the counted emails that are categorized
        df = pd.DataFrame(counting_dict.items(), columns=['Item', 'Count'])
    
        title_str = "Counts (" + "Start: " + date_start_str + " | End: " + date_end_str + ")"

        #Removing the axis for matplotlib and creating a visual table of the counted emails that are categorize
        fig, ax = plt.subplots()
        plt.title(title_str)
        ax.axis('off')
        ax.axis('tight')
        table = ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
        table.scale(1, 2)
        plt.savefig(COUNTING_IMAGE_FILE_NAME)  # saves plot locally,
        
        #prints a tabulate table using the pandas dataframe
        print("\n")
        print(title_str)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex='never'))
        
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

#Generates categories visualizations
def generate_categories_viz():
    try:
        
        if os.stat(CATEGORIES_DATA_FILE_NAME).st_size == 0:
            return
        
        #Pandas dataframe for the counted emails that are categorized
        df = pd.read_csv(CATEGORIES_DATA_FILE_NAME, sep = "\t")
    
        #Removing the axis for matplotlib and creating a visual table of the counted emails that are categorize
        fig, ax = plt.subplots()
        plt.title('Number of Email(s) Categories')
        ax.axis('off')
        ax.axis('tight')
        table = ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
        table.scale(1, 2)
        plt.savefig(CATEGORIES_IMAGE_FILE_NAME)  # saves plot
        
        #prints a tabulate table using the pandas dataframe
        print("\n")
        print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex='never'))
        
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

#Generates flagged email visualization
def generate_generic_viz(flagged_counter_int,email_data_file,email_image_file,title):
    if (flagged_counter_int  > 0):
        try:

            # Create a dataframe from the list of dictionaries
            df = pd.read_csv(email_data_file, sep = "\t")

            #Removing the axis for matplotlib and creating a visual table of the counted emails that are categorize
            fig, ax = plt.subplots()
            plt.title(title)
            ax.axis('off')
            ax.axis('tight')
            table = ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
            table.scale(2, 2)
            plt.savefig(email_image_file, dpi=150)  # saves plot

        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

        try:

            print("\n")
            print(title)
            print(tabulate(df, headers = 'keys', tablefmt = 'fancy_grid',showindex='never'))

        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
            
            #Prints some ouputs to Command Line
    print("\n")
    print("Number of Flagged emails: ", flagged_counter_int)

#Generates word cloud visualization
def generate_word_cloud_viz():
    try:
        wc_cleaned_content_file = open(WORD_CLOUD_CLEANED_FILE_NAME, "r", encoding = "utf-8").read()
        
        #Sets stopwords  for cloud
        stop_words = ["said", "email", "s", "will", "u", "re", "3A", "2F", "safelinks", "reserved", "https"] + list(STOPWORDS) #customized stopword list
        
        #Generates word cloud
        word_cloud = WordCloud(stopwords = stop_words).generate(str(wc_cleaned_content_file))
        plt.clf()
        plt.imshow(word_cloud)
        plt.axis('off')
        plt.savefig(WORD_CLOUD_IMAGE_FILE_NAME)
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

#Function to remove converse characters ("\u202a") and pop directional formatting characters from strings ("\u202c")
#Code borrowed from https://stackoverflow.com/questions/49267999/remove-u202a-from-python-string
def cleanup(inp):
    new_char = ""
    for char in inp:
        if char not in ["\u202a", "\u202c"]:
            new_char += char
    return new_char


#Removes hyperlink information from email body text to make more meaningful clouds
def word_cloud_content_clean():
    try:
        #Opens extracted email body text and cleansed text storage file
        wc_content= open(WORD_CLOUD_FILE_NAME, "r", encoding = "utf-8").read()
        wc_content_cleaned = open(WORD_CLOUD_CLEANED_FILE_NAME, "w+", encoding = "utf-8")
        
        #Create indices counter variable
        indices_counter_int = 0
        
        #Assigns variables to tags preceeding link information
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



#Identifiies unique elements within a list
def unique (list1):
    unique_elements_list = []
    for item in list1:
        if item not in unique_elements_list:
            unique_elements_list.append(item)
    return unique_elements_list


def main():  
    
    print("\n")
    print("Welcome to Outlook Analyzer!")


    #Receives and checks max numbebr of emails to extract
    while True:
        max_email_number_to_extract_input = input("Max number of email messages you would like to extract (between 50 and 100000)? (hit Enter for default: 500)") or 500

        try:
            int(max_email_number_to_extract_input)
            if int(max_email_number_to_extract_input) >= 50 and int(max_email_number_to_extract_input) <= 100000:
                break;
        except ValueError:
            print("Please enter a valid integer between 50 and 100000.")

    #Receives and checks user input for oldest email cut-off date
    while True:
        date_start_input = input("From how far back would you like to collect and analyze emails in months or days (e.g. 10m, 12d)? (Hit enter for default: 12 months ago)") or "12m"

        try:
            int(date_start_input[0:-1])
            if date_start_input[-1] == "m" or date_start_input[-1] == "d":
                break;
        except ValueError:
            print("This is not a valid format. Please enter as '##m' or '##d' where d is for days and m is for months  (e.g. 10d or 1m")

    #Receives and checks user input for email recency cut-off date
    while True:
        date_end_input = input("What's the cutoff for the most recent emails you'd like to collect and analyze in months or days (e.g. 1m, 10d)? (Hit enter for default: today)") or "0m"

        try:
            int(date_end_input[0:-1])
            if date_end_input[-1] == "m" or date_end_input[-1] == "d":
                break;
        except ValueError:
            print("This is not a valid format. Please enter as '##m' or '##d' where d is for days and m is for months  (e.g. 10d or 1m")

    extract_outlook_information(max_email_number_to_extract_input,date_start_input,date_end_input)

    #Displays any errors generated throughout program execution
    if ERROR_LIST != []:
        print("\n")
        print("Some errors occurred during execution:")
        for item in ERROR_LIST:
            print(item)

if __name__ == "__main__":
    main()
