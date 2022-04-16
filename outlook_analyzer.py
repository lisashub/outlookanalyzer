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
from dateutil.relativedelta import relativedelta
import sys

ERROR_LIST = []
TIME_STR = time.strftime("%Y%m%d-%H%M%S")
TEMP_DIR = "C:\WINDOWS\Temp"

WORD_CLOUD_CLEANED_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud_text_cleaned.txt"
WORD_CLOUD_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud_text.txt"
WORD_CLOUD_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "word_cloud.jpg"
UNREAD_SENDERS_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "unread_senders.txt"
CATEGORIES_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "categories.txt"
FLAGGED_EMAIL_DATA_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "flagged_email.txt"
FLAGGED_EMAIL_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "flagged_email_list.png"
SENDER_PLOT_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "sender_plot.jpg"
CATEGORIES_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "categories.jpg"


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
    
    #additional variable creation
    unread_senders_raw_list = [] #list variable to capture unread email senders with dupes
    unread_senders_unique_dict = {} #dictionary variable to capture unique unread senders with counts
    categories_senders_list = [] #dictionary variable to capture category and sender information
    category_list = [] #list variable to capture email category 
    category_dict = {} #dictionary variable to capture email category
    flagged_messages_list = [] #list to capture flagged messege info
    
    categories_counter_int = 0
    flagged_counter_int = 0
    message_counter_int = 0

    messages.Sort("[ReceivedTime]",True)

    # Setup end_date for month or days for date range filter
    if date_end_input[-1] == "m":
        month_int = int(date_end_input[0:-1])

        if month_int == 0:
            end_date = date.today()
        else:
           end_date = date.today() - relativedelta(months=+month_int)

    elif date_end_input[-1] == "d":
        day_int = int(date_end_input[0:-1])
 
        if day_int == 0:
            end_date = date.today()
        else:
            end_date = date.today() - timedelta(days=day_int)

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
    
    for item in tqdm(filtered_message):

        #check and store unread email info
        try:
            if (item.Unread == True):
                sender = item.SenderEmailAddress
                unread_senders_raw_list.append(sender)
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
        
       #check and store categories info
        try:
            if item.Categories:
                item_categories = item.Categories.split(",")
                for category in item_categories:
                    category_list.append(category)
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

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
            subject = task.subject
            # Remove invisible white space / pointers that pandas cannot handle
            clean_subject = cleanup(subject)           
            subject = [clean_subject.encode("utf-8").strip()]
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

       
        message_counter_int = message_counter_int + 1

        # Check if max number of email has been reached       
        if message_counter_int >= int(max_email_number_to_extract_input):
            break

    unread_senders_data_gen(unread_senders_raw_list, unread_senders_unique_dict,sender_data_file)
    generate_unread_senders_viz()
    category_data_gen(categories_counter_int, category_list, category_dict, categories_data_file)
    generate_categories_viz()
    flagged_email_data_gen(flagged_messages_list, flagged_email_data_file)
    generate_flagged_viz(flagged_counter_int)
    word_cloud_extract(messages)
    word_cloud_display()

    # ??? Is needed to close the files again here when they are also being closed in their respective functions?
    sender_data_file.close()
    categories_data_file.close()
    
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

def flagged_email_data_gen(flagged_messages_list, flagged_email_data_file):

    # Have to generate a text file to decode the utf8 data
    try:
        print("Subject","\t",'SenderEmailAddress',"\t", 'ReceivedTime', file = flagged_email_data_file)

        for item in flagged_messages_list:

            if item['Class'] == 43: 
                print('\n'.join(s.decode('utf-8', 'ignore') for s in item['subject']),"\t",item['SenderEmailAddress'], "\t", item['ReceivedTime'], file = flagged_email_data_file)
            else:
                print('\n'.join(s.decode('utf-8', 'ignore') for s in item['subject']),"\t", "-", "\t", "-", file = flagged_email_data_file)

        flagged_email_data_file.close()
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))


def category_data_gen(categories_counter_int,category_list,category_dict,categories_data_file):
    try:
        unique_categories = unique(category_list) #sends category list to function named "unique" and saves list of unique values to variable
        for category in unique_categories: #loops through unique categories and counts occurrances; saves results into category_dict
            category_dict[category] = category_list.count(category)
        
        for item in category_dict.items():
            print(item[0], "\t", item[1], file = categories_data_file)
            categories_counter_int = categories_counter_int + 1
            
        print('\n')
        print('Number of categories:', len(unique_categories))  # print total number or emails categorize
            
        categories_data_file.close()
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
    
#Generates categories visualizations
def generate_categories_viz():
    try:
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

# Function to remove certain utf8 characters from strings
def cleanup(inp):
    new_char = ""
    for char in inp:
        if char not in ["\u202a", "\u202c"]:
            new_char += char
    return new_char

#Generates flagged email visualization
def generate_flagged_viz(flagged_counter_int):
    if (flagged_counter_int  > 0):
        try:

            # Create a dataframe from the list of dictionaries
            df = pd.read_csv(FLAGGED_EMAIL_DATA_FILE_NAME, sep = "\t")


            #Removing the axis for matplotlib and creating a visual table of the counted emails that are categorize
            fig, ax = plt.subplots()
            plt.title('Flagged Email(s) / Todo')
            ax.axis('off')
            ax.axis('tight')
            table = ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
            table.scale(2, 2)
            plt.savefig(FLAGGED_EMAIL_IMAGE_FILE_NAME, dpi=150)  # saves plot

            # Old method whichs works but not exactly what I was looking for
            # dfi.export(df, FLAGGED_EMAIL_IMAGE_FILE_NAME)

        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

        try:

            print("\n")
            print("Flagged Emails and To Do Items")
            print(tabulate(df, headers = 'keys', tablefmt = 'fancy_grid',showindex='never'))

        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
            
            #Prints some ouputs to Command Line
    print("\n")
    print("Number of Flagged emails: ", flagged_counter_int)


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
