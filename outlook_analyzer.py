import win32com.client #core extraction library
from tqdm import tqdm # library to display extraction progress bar
import pandas as pd # library to tabulate data and generate plot/image
from tabulate import tabulate
import matplotlib.pyplot as plt
import dataframe_image as dfi
from wordcloud import WordCloud, STOPWORDS #word cloud generation library
import re #regex library used to clean data
import time
from datetime import timedelta
from datetime import date
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys
from matplotlib.backends.backend_pdf import PdfPages
import random
import string
from fpdf import FPDF
from glob import glob
from PyPDF2 import PdfFileMerger
import os
import subprocess
import time


ERROR_LIST = []
TIME_STR = time.strftime("%Y%m%d-%H%M%S")
TEMP_FOLDER = "outlookanalyzer"
TEMP_DIR = "C:\\WINDOWS\\Temp" + "\\" + TEMP_FOLDER 
FINAL_REPORT = "C:\\WINDOWS\\Temp\\" + TIME_STR + "_outlook_analyzer_report.pdf" 

# Check whether the specified path exists or not
is_temp_dir_exist = os.path.exists(TEMP_DIR)

if not is_temp_dir_exist:
  
  # Create a new directory because it does not exist 
  os.makedirs(TEMP_DIR)
  print("The new directory is created!")

WORD_CLOUD_CLEANED_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud_text_cleaned.txt"
WORD_CLOUD_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud_text.txt"
WORD_CLOUD_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud.jpg"
UNREAD_SENDERS_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "unread_senders.txt"
CATEGORIES_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "categories.txt"
FLAGGED_EMAIL_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "flagged_email.txt"
IMPORTANT_EMAIL_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "important_email.txt"
FLAGGED_EMAIL_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "flagged_email_list.png"
IMPORTANT_EMAIL_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "important_email_list.png"
SENDER_PLOT_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "sender_plot.jpg"
CATEGORIES_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "categories.jpg"
COUNTING_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "counting.jpg"

image_file_name_dict = {'icon'   : {"image_path": "icon.png", "x":"0", "y":"0", "w":"35", "h":"30"},
             'blue'   : {"image_path": "true_blue.jpg", "x":"35", "y":"0", "w":"175", "h":"30"},
             'word_cloud' : {"image_path": WORD_CLOUD_IMAGE_FILE_NAME, "x":"0", "y":"75", "w":"300", "h":"150"},
             'sender_plot'  :{"image_path": SENDER_PLOT_FILE_NAME, "x":"0", "y":"150", "w":"210", "h":"100"} }

#             'count_overview' :{"image_path": COUNTING_IMAGE_FILE_NAME, "x":"0", "y":"130", "w":"175", "h":"30"},

def append_to_error_list(function_name, error_text):
    ERROR_LIST.append("function: " + function_name + " | " +  "error: " + error_text)

#Extracts data from Outlook
def extract_outlook_information(max_email_number_to_extract_input,date_start_input,date_end_input): #to modify as new features required

    #Connection to Outlook object model established
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
    inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
    messages = inbox.Items #variable for items in inbox

    #create data files

    sender_data_file = open(UNREAD_SENDERS_DATA_FILE_NAME, "w+", encoding = "utf-8")
    categories_data_file = open(CATEGORIES_DATA_FILE_NAME, "w+", encoding = "utf-8")
    flagged_email_data_file = open(FLAGGED_EMAIL_DATA_FILE_NAME, "w+", encoding = "utf-8")
    important_email_data_file = open(IMPORTANT_EMAIL_DATA_FILE_NAME, "w+", encoding = "utf-8")
    
    #additional variable creation
    unread_senders_raw_list = [] #list variable to capture unread email senders with dupes
    unread_senders_unique_dict = {} #dictionary variable to capture unique unread senders with counts
    category_list = [] #list variable to capture email category 
    flagged_messages_list = [] #list to capture flagged messege info
    important_messages_list = [] #list to capture important messege info
    counting_dict = {} #dictionary variable to capture different counts

    categories_counter_int = 0
    number_of_times_categories_assigned_counter_int = 0
    flagged_counter_int = 0
    message_counter_int = 0
    message_read_counter_int = 0
    message_unread_counter_int = 0
    important_count_int = 0

    messages.Sort("[ReceivedTime]",True)

    # Setup end_date for month or days for date range filter
    if date_end_input[-1] == "m":
        month_int = int(date_end_input[0:-1])

        if month_int == 0:
            # end_date = date.today()
            end_date = datetime.now()
        else:
           end_date = datetime.now() - relativedelta(months=+month_int)

    elif date_end_input[-1] == "d":
        day_int = int(date_end_input[0:-1])
 
        if day_int == 0:
            end_date = datetime.now()
        else:
            end_date = datetime.now() - timedelta(days=day_int)

    # Setup start_date for month or days for date range filter
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

    # Convert time into string format for filtering messages
    date_start_str = start_date.strftime('%m/%d/%Y %H:%M %p')
    date_end_str = end_date.strftime('%m/%d/%Y %H:%M %p')

    filtered_message = messages.Restrict("[ReceivedTime] >= '" + date_start_str + "' AND [ReceivedTime] <= '" + date_end_str + "'")
    
    print("Extracting email messages:")
    for item in tqdm(filtered_message):

        #check and store unread email info
        try:
            if (item.UnRead == True):

                message_unread_counter_int = message_unread_counter_int + 1

                if item.Class == 43:
                    if item.SenderEmailType == "EX":
                        sender = item.Sender.GetExchangeUser().PrimarySmtpAddress
                    else:
                        sender = item.SenderEmailAddress
                else:
                    sender = item.SenderEmailAddress
    
                unread_senders_raw_list.append(sender)

            else:
                message_read_counter_int = message_read_counter_int + 1

        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
        
       #check and store categories info
        try:
            if item.Categories:
                item_categories = item.Categories.split(",")
                categories_counter_int = categories_counter_int + 1
                for category in item_categories:
                    category_list.append(category.strip())
                    number_of_times_categories_assigned_counter_int = number_of_times_categories_assigned_counter_int + 1
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

        #check and store emails marked as important
        # 2 means is an email is marked as important
        # 43 is the class for regular email
        if (item.Importance == 2 and item.Class == 43):
            important_messages_dict = {} #dict to capture important message info (create a dict for each email)

            try:
                subject = item.Subject
                # Remove invisible white space / pointers that pandas cannot handle
                clean_subject = cleanup(subject)
                subject = [clean_subject.encode("utf-8").strip()] # Common for subjects to have emoji or utf8 data so need encode as utf8
                important_messages_dict['subject'] = subject
            except Exception as e:
                append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
        
            try:
                email_class = item.Class
                important_messages_dict['Class'] = email_class
            except Exception as e:
                append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

            try:
                received_time = item.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S")
                important_messages_dict['ReceivedTime'] = received_time
            except Exception as e:
                append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

            # Do not want the long exchange details, for "EX" SenderEmailType, get the PrimarySmtpAddress
            if item.SenderEmailType == "EX":
                try:
                    sender_email = item.Sender.GetExchangeUser().PrimarySmtpAddress
                except Exception as e:
                    append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
            else:
                try:
                    sender_email = item.SenderEmailAddress
                except Exception as e:
                    append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

            important_messages_dict['SenderEmailAddress'] = sender_email

            # Add the attributes for each item to the important_messages_dict
            important_messages_list.append(important_messages_dict)

            # Keep track of the numbers of important items
            important_count_int = important_count_int + 1

        message_counter_int = message_counter_int + 1

        # Check if max number of email has been reached       
        if message_counter_int >= int(max_email_number_to_extract_input):
            break

    todo_folder = outlook.GetDefaultFolder(28) #outlook.GetDefaultFolder(28) is for the todo/flagged items

    todo_items = todo_folder.Items
    tasks = todo_items.Restrict("[Complete] = FALSE")

    print("Extracting tasks/flagged items:")
    for task in tqdm(tasks):      

        #check and store flagged email/tasks/todo

        flagged_messages_dict = {} #dict to capture flagged message info (create a dict for each task)

        # Assign value and add to dict for several attributes

        # Tasks that do not come in as email only have subject
        try:
            subject = task.Subject
            # Remove invisible white space / pointers that pandas cannot handle
            clean_subject = cleanup(subject)           
            subject = [clean_subject.encode("utf-8").strip()] # Common for subjects to have emoji or utf8 data so need encode as utf8
            flagged_messages_dict['subject'] = subject
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

        try:
            item_class = task.Class
            flagged_messages_dict['Class'] = item_class
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

        # class 43 is standard MailItem. ReportItem/MeetingItem are a different class.
        if task.Class == 43:
        
            try:
                received_time = task.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S")
                flagged_messages_dict['ReceivedTime'] = received_time
            except Exception as e:
                append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

            # Do not want the long exchange details, for "EX" SenderEmailType, get the PrimarySmtpAddress
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

    unread_senders_data_gen(unread_senders_raw_list, unread_senders_unique_dict,sender_data_file)
    generate_unread_senders_viz()
    category_data_gen(category_list, categories_data_file)

    # Converts collected data into a text file for flagged messages 
    build_text_with_subject_senderemail_receivedtime(flagged_messages_list, flagged_email_data_file)
    # Converts collected data into a text file for flagged messages 
    build_text_with_subject_senderemail_receivedtime(important_messages_list, important_email_data_file)

    count_title_str = "Counts (" + "Start: " + date_start_str + " | End: " + date_end_str + ")"
    figure_column_list = ['Item', 'Count']
    convert_dict_to_df_to_figure_to_pdf(counting_dict, count_title_str,figure_column_list)

    figure_column_list = ['Category Name', 'Count']
    convert_csv_to_df_to_figure_to_pdf(CATEGORIES_DATA_FILE_NAME,"Categories",figure_column_list)

    figure_column_list = ["Subject","Sender Email","Date"]
    convert_csv_to_df_to_figure_to_pdf(FLAGGED_EMAIL_DATA_FILE_NAME,"Flagged email / Todo", figure_column_list)

    figure_column_list = ["Subject","Sender Email","Date"]
    convert_csv_to_df_to_figure_to_pdf(IMPORTANT_EMAIL_DATA_FILE_NAME,"Email sent as Important",figure_column_list)

    word_cloud_extract(messages)
    word_cloud_display()
    
    create_pdf_cover_page()
    pdf_merge()

#Extracts word cloud information from messages
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

#Generates data for unread senders visualizations
def unread_senders_data_gen(unread_senders_raw_list,unread_senders_unique_dict,sender_data_file):
    try:
        unread_senders_unique_list = unique(unread_senders_raw_list)
        
        for sender in unread_senders_unique_list:
            unread_senders_unique_dict[sender] = unread_senders_raw_list.count(sender)
        
        unread_senders_unique_list = sorted(unread_senders_unique_dict.items(), key = lambda x:x[1], reverse = True) # sorts dict and stores as list
        unread_senders_unique_sorted_dict = dict(unread_senders_unique_list) #converts list back to dict
    
        senders_counter_int = 0 # counter variable to help count collected senders
        for item in unread_senders_unique_sorted_dict.items(): #write's top nth senders and email count to file
            if (senders_counter_int<10):
                print(item[0], "\t", item[1], file = sender_data_file)
                senders_counter_int = senders_counter_int + 1
        
        sender_data_file.close()
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
    
#Generates unread senders visualizations
def generate_unread_senders_viz():
    
    #Reads unread sender data and generates visualization; saves and displays
    try:
        sender_table = pd.read_table(UNREAD_SENDERS_DATA_FILE_NAME, sep = '\t', header = None)
        plot = sender_table.groupby([0]).sum().plot(kind='pie', y=1, labeldistance=None, autopct='%1.0f%%', title="Senders of Unread Emails")
        plot.legend(bbox_to_anchor=(1,1)) #Sets legend details
        plot.set_ylabel("Senders") #Set label detail
        plot.figure.savefig(SENDER_PLOT_FILE_NAME, bbox_inches='tight') #saves plot locally
        print("\n","Top 10 Senders of Unread Emails: ", "\n", sender_table)
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

def build_text_with_subject_senderemail_receivedtime(messages_list, email_data_file):

    # Have to generate a text file to decode the utf8 data
    try:

        for item in messages_list:

            if item['Class'] == 43: 
                print('\n'.join(s.decode('utf-8', 'ignore') for s in item['subject']),"\t",item['SenderEmailAddress'], "\t", item['ReceivedTime'], file = email_data_file)
            else:
                print('\n'.join(s.decode('utf-8', 'ignore') for s in item['subject']),"\t", "-", "\t", "-", file = email_data_file)

        email_data_file.close()
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

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
def convert_dict_to_df_to_figure_to_pdf(counting_dict, title_str, columns_list):
    try:
        #Pandas dataframe for the counted emails that are categorized
        df = pd.DataFrame(counting_dict.items(), columns=columns_list)   

        fig, ax =plt.subplots()
        plt.title(title_str)
        ax.axis('tight')
        ax.axis('off')
        table = ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
        table.scale(1, 2)
        # plt.savefig(COUNTING_IMAGE_FILE_NAME)  # saves plot locally,

        random_string = string.ascii_lowercase
        letters = string.ascii_lowercase
        random_string = "b_" + ( ''.join(random.choice(letters) for i in range(10)) ) + ".pdf"
        file_path = TEMP_DIR + "\\" + random_string

        pp = PdfPages(file_path)
        pp.savefig(fig, bbox_inches='tight')

        pp.close()
        
        #prints a tabulate table using the pandas dataframe
        print("\n")
        print(title_str)
        print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex='never'))
        
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

# Function to remove certain utf8 characters from strings
def cleanup(inp):
    new_char = ""
    for char in inp:
        if char not in ["\u202a", "\u202c"]:
            new_char += char
    return new_char

def pdf_merge():
    ''' Merges all the pdf files in current directory '''
    merger = PdfFileMerger()
    location_to_check_str = TEMP_DIR + "\\" + "*.pdf"
    allpdfs = [a for a in glob(location_to_check_str)]
    [merger.append(pdf) for pdf in allpdfs]
    with open(FINAL_REPORT, "wb") as new_file:
        merger.write(new_file)

    subprocess.Popen([FINAL_REPORT],shell=True)

def create_pdf_cover_page():
    try:
    
        # Code for creating PDF dashboard
        pdf = FPDF()
        pdf.add_page()  # adds pdf page
        pdf.set_font('Arial', 'B', 18)  # sets pdf fonts
        pdf.cell(0, 60, 'Outlook Analzyer Dashboard', 0, 0, align='C')  # Puts in title
        # pdf.cell(-190, 75, date, 0, 0, align='C') # Puts in real time date

        for image_id, image_info in image_file_name_dict.items():
            print("\nItem:", image_id)
            
            for key in image_info:
                print(key + ':', image_info[key])
                pdf.image(image_info['image_path'], x=int(image_info['x']), y=int(image_info['y']), w=int(image_info['w']), h=int(image_info['h']))  

        file_path = TEMP_DIR + "\\" + "a_cover_page.pdf"

        pdf.output(file_path, 'F')  # saves pdf into local file
        pdf.open()


    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))



def convert_csv_to_df_to_figure_to_pdf(email_data_file,title_str,columns_list):

    try:
        df = pd.read_csv(email_data_file, sep = "\t", encoding ='utf-8', names=columns_list)

        fig, ax =plt.subplots(figsize=(12,4))
        plt.title(title_str)
        ax.axis('tight')
        ax.axis('off')
        table = ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
        table.scale(1, 2)

        random_string = string.ascii_lowercase
        letters = string.ascii_lowercase
        random_string = "c_" + ( ''.join(random.choice(letters) for i in range(10)) ) + ".pdf"

        file_path = TEMP_DIR + "\\" + random_string

        pp = PdfPages(file_path)
        pp.savefig(fig, bbox_inches='tight')

        pp.close()
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

    try:

        print("\n")
        print(tabulate(df, headers = 'keys', tablefmt = 'fancy_grid',showindex='never'))

    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

#Removes hyperlink information from email body to produce more meaningful clouds
def word_cloud_content_clean():
    
    try:
        wc_content= open(WORD_CLOUD_FILE_NAME, "r", encoding = "utf-8").read()
        wc_content_cleaned = open(WORD_CLOUD_CLEANED_FILE_NAME, "w+", encoding = "utf-8")
        
        sub1 = '<http'
        sub2 = '<mail'
        
        indices1_link = [m.start() for m in re.finditer(sub1, wc_content)]
        indices2_mail = [m.start() for m in re.finditer(sub2, wc_content)]
        indices1_link.extend(indices2_mail)
        indices1_link.sort()
        
        indices2 = []
        
        for indices in indices1_link:
                end_indices = wc_content[indices:].find('>')
                indices2.append(end_indices+indices)
        
        i = 0
        while i < len(indices2)-1:
            print(wc_content[indices2[i]+1:indices1_link[i+1]], file = wc_content_cleaned)
            i = i + 1
            
        wc_content_cleaned.close()
    
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

#Generates word cloud from cleaned email text
def word_cloud_generate():
    wc_content= open(WORD_CLOUD_CLEANED_FILE_NAME, "r", encoding = "utf-8").read()
    stop_words = ["said", "email", "s", "will", "u", "re", "3A", "2F", "safelinks", "reserved", "https"] + list(STOPWORDS) #customized stopword list
    wordcloud = WordCloud(stopwords = stop_words).generate(str(wc_content))
    return(wordcloud)

#Saves and displays word cloud to user
def word_cloud_display():
    try:
        plt.clf()
        plt.imshow(word_cloud_generate())
        plt.axis('off')
        plt.savefig(WORD_CLOUD_IMAGE_FILE_NAME)
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

#Identifiies unique elements within a list
def unique (list1):
    unique_elements_list = []
    for item in list1:
        if item not in unique_elements_list:
            unique_elements_list.append(item)
    return unique_elements_list


def delete_stuff ():
    dir_list = os.listdir(TEMP_DIR) #returns the list of all files and directories in the specified path

    print(dir_list)
    # Goes through directory and removes/cleans the files with specified extensions (i.e .txt, .tmp, .png, etc.)
    while True:
            # dir_option = input('Would you like to clean your directory? (y/n):')
            # if dir_option == 'y' or dir_option == 'Y':
                for item in dir_list:
                    if item.endswith('.txt') or item.endswith('.png') or item.endswith('.jpg') or item.endswith('.pdf'): # comment this out just in case you don't want to remove the images in the PDF code above
                        os.remove(os.path.join(TEMP_DIR, item))
                print('Directory cleaned')
                break;
            # elif dir_option == 'n' or dir_option == 'N':
            #     pass
            #     break;


#UI code; checks whether new extraction required and calls if necessary
def main():  
    
    print("\n")
    print("Welcome to Outlook Analyzer!")

    user_input = input("Would you like to extract fresh data from Outlook? (Y/N)")

    #Logic to determine whether extraction is required
    if user_input == "Y" or user_input == "y":

        # Check if user provided an actual integer
        while True:
            max_email_number_to_extract_input = input("Max number of email messages you would like to extract (between 50 and 100000)? (hit Enter for default: 500)") or 500

            try:
                int(max_email_number_to_extract_input)
                if int(max_email_number_to_extract_input) >= 50 and int(max_email_number_to_extract_input) <= 100000:
                    break;
            except ValueError:
                print("Please enter a valid integer between 50 and 100000.")

        # Check if user entered in proper format for day/month
        while True:
            date_start_input = input("From how far back would you like to collect and analyze emails in months or days (e.g. 10m, 12d)? (Hit enter for default: 12 months ago)") or "12m"

            try:
                int(date_start_input[0:-1])
                if date_start_input[-1] == "m" or date_start_input[-1] == "d":
                    break;
            except ValueError:
                print("This is not a valid format. Please enter as '##m' or '##d' where d is for days and m is for months  (e.g. 10d or 1m")

        # Check if user entered in proper format for day/month
        while True:
            date_end_input = input("What's the cutoff for the most recent emails you'd like to collect and analyze in months or days (e.g. 1m, 10d)? (Hit enter for default: today)") or "0m"

            try:
                int(date_end_input[0:-1])
                if date_end_input[-1] == "m" or date_end_input[-1] == "d":
                    break;
            except ValueError:
                print("This is not a valid format. Please enter as '##m' or '##d' where d is for days and m is for months  (e.g. 10d or 1m")

        delete_stuff()
        extract_outlook_information(max_email_number_to_extract_input,date_start_input,date_end_input)

    else:
        print("Sorry, the provided response was not understood.")

    #displays errors generated throughout program execution
    if ERROR_LIST != []:
        print("\n")
        print("Some errors occurred during execution:")
        for item in ERROR_LIST:
            print(item)

if __name__ == "__main__":
    main()
