
import win32com.client #core extraction library
from tqdm import tqdm # library to display extraction progress bar
import pandas as pd # library to tabulate data and generate plot/image
from tabulate import tabulate
import matplotlib.pyplot as plt
import dataframe_image as dfi
from wordcloud import WordCloud, STOPWORDS #word cloud generation library
import re #regex library used to clean data
import json
from datetime import timedelta
from datetime import date
from dateutil.relativedelta import relativedelta
import sys

ERROR_LIST = []
TEMP_DIR = "C:\WINDOWS\Temp"
EMAIL_OUTPUT = "email_out"
EMAIL_FILE_PATH = TEMP_DIR + "\\" + EMAIL_OUTPUT
WORD_CLOUD_CLEANED_FILE_NAME = TEMP_DIR + "\\" + "word_cloud_text_cleaned.txt"
WORD_CLOUD_FILE_NAME = TEMP_DIR + "\\" + "word_cloud_text.txt"
WORD_CLOUD_IMAGE_FILE_NAME = TEMP_DIR + "\\" + "word_cloud.jpg"

def append_to_error_list(function_name, error_text):
    ERROR_LIST.append("function: " + function_name + " | " +  "error: " + error_text)

def read_email(save_data_boolean,data_source_str):

    if data_source_str == "Outlook":

        print("Connecting to Outlook...")

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
        inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
        messages = inbox.Items #variable for items in inbox

        # Will be adding this stuff in later...

        # # Example using timedelta by days
        # # received_dt = date.today() - timedelta(days=30)
        # # e_dt = date.today() - timedelta(days=30)
        # # s_dt = date.today() - timedelta(days=60)

        # # Example using relativedelta by month
        # end_date = date.today() - relativedelta(months=+1)
        # start_date = date.today() - relativedelta(months=+2)

        # start_date_str = start_date.strftime('%m/%d/%Y %H:%M %p')
        # end_date_str = end_date.strftime('%m/%d/%Y %H:%M %p')

        # print("Extracting email in this date range: ")
        # print("Staring date: " + str(start_date))
        # print("Ending date: " + str(end_date))

        # # Other ways to restrict email by ReceivedTime
        # # received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
        # # filtered_message = messages.Restrict("[ReceivedTime] >= '" + str(received_dt) + "'")

        # filtered_message = messages.Restrict("[ReceivedTime] >= '" + start_date_str + "' AND [ReceivedTime] <= '" + end_date_str + "'")

        # # today = datetime.date.today()
        # # first = today.replace(day=1)
        # # lastMonth = first - datetime.timedelta(days=1)
        # # print(lastMonth.strftime("%Y%m"))

        # # Can also restrict emails in other ways
        # # current_message = messages.Restrict("[SenderEmailAddress] = 'notifications@github.com' ")

        # # Get the most recent email
        # # current_message = messages.GetLast()
        # # print(current_message)

        all_messages_list = []

        messages.Sort("[ReceivedTime]",True)
        for item in tqdm(messages):
        # for item in tqdm(filtered_message):
            messages_dict = {}

            # Some attributes are only applicable to email, not other items such as meeting invites
            if item.Class == 43: # normal email are class 43 (meeting invites are 53, responses are 56 )

                try:
                    alternate_recipient_allowed = item.AlternateRecipientAllowed
                    messages_dict['AlternateRecipientAllowed'] = alternate_recipient_allowed
                except Exception as e:
                    print("error:" + str(e))

                try:
                    is_marked_as_task = item.IsMarkedAsTask
                    messages_dict['IsMarkedAsTask'] = is_marked_as_task
                except Exception as e:
                    print("error:" + str(e))

                try:
                    read_receipt_requested = item.ReadReceiptRequested
                    messages_dict['ReadReceiptRequested'] = read_receipt_requested
                except Exception as e:
                    print("error:" + str(e))

                try:
                    to = item.To
                    messages_dict['To'] = to
                except Exception as e:
                    print("error:" + str(e))

                try:
                    cc = item.CC
                    messages_dict['CC'] = cc
                except Exception as e:
                    print("error:" + str(e))

                try:
                    bcc = item.BCC
                    messages_dict['BCC'] = bcc
                except Exception as e:
                    print("error:" + str(e))

            # Continue with other attributes applicable to email and other items such as meeting invites

            try:
                attachments_list = []
                # Loops over each attachment and adds a list of attachment file names
                for attachment in item.Attachments:
                    attachments_list.append(attachment.FileName)
                messages_dict['Attachments'] = attachments_list
            except Exception as e:
                print("error:" + str(e))

            try:
                auto_forwarded = item.AutoForwarded
                messages_dict['AutoForwarded'] = auto_forwarded
            except Exception as e:
                print("error:" + str(e))

            try:
                email_class = item.Class
                messages_dict['Class'] = email_class
            except Exception as e:
                print("error:" + str(e))

            try:
                # Split the CSV string from Categories into a list 
                messages_dict['Categories'] = item.Categories.split(",")
            except Exception as e:
                print("error:" + str(e))

            try:
                companies = item.Companies
                messages_dict['Companies'] = companies
            except Exception as e:
                print("error:" + str(e))

            try:
                importance = item.Importance
                messages_dict['Importance'] = importance
            except Exception as e:
                print("error:" + str(e))

            try:
                sender_name = item.SenderName
                messages_dict['SenderName'] = sender_name
            except Exception as e:
                print("error:" + str(e))

            try:
                sender_email_address = item.SenderEmailAddress
                messages_dict['SenderEmailAddress'] = sender_email_address
            except Exception as e:
                print("error:" + str(e))

            try:
                sent_on = item.SentOn
                messages_dict['SentOn'] = sent_on
            except Exception as e:
                print("error:" + str(e))

            try:
                sender_email_type = item.SenderEmailType
                messages_dict['SenderEmailType'] = sender_email_type
            except Exception as e:
                print("error:" + str(e))

            try:
                reminder_set = item.ReminderSet
                messages_dict['ReminderSet'] = reminder_set
            except Exception as e:
                print("error:" + str(e))

            try:
                recipients_list = []
                # Loops over each recipient and adds a list of recipient  names
                for recipient in item.Recipients:
                    recipients_list.append(recipient.Name)
                messages_dict['Recipients'] = recipients_list
            except Exception as e:
                print("error:" + str(e))

            try:
                sensitivity = item.Sensitivity
                messages_dict['Sensitivity'] = sensitivity
            except Exception as e:
                print("error:" + str(e))

            try:
                flag_request = item.FlagRequest
                messages_dict['FlagRequest'] = flag_request
            except Exception as e:
                print("error:" + str(e))

            try:
                unread = item.UnRead
                messages_dict['UnRead'] = unread
            except Exception as e:
                print("error:" + str(e))

            try:
                size = item.Size
                messages_dict['Size'] = size
            except Exception as e:
                print("error:" + str(e))

            try:
                subject = item.Subject
                messages_dict['Subject'] = subject
            except Exception as e:
                print("error:" + str(e))
            
            try:
                body = item.Body
                messages_dict['Body'] = body
            except Exception as e:
                print("error:" + str(e))

            try:
                received_time = item.ReceivedTime
                # received_time = item.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S")
                messages_dict['ReceivedTime'] = received_time
            except Exception as e:
                print("error:" + str(e))

            all_messages_list.append(messages_dict)

        # Save email data to disk as a list of JSON key value items
        if save_data_boolean:
            print("Saving Outlook email data to: " + EMAIL_FILE_PATH)
            json_object = json.dumps(all_messages_list, indent = 4, sort_keys=True, default=str)

            with open(EMAIL_FILE_PATH, "w+") as outfile:
                outfile.write(json_object)
                    
    elif data_source_str == "JSON":

        # Opening JSON file
        with open(EMAIL_FILE_PATH) as json_file:
            all_messages_list = json.load(json_file)
        
    return (all_messages_list)

#Extracts data from Outlook
def extract_email_information_from_messages_list(all_messages_list): #to modify as new features required
    
    #Connection to Outlook object model established
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
    inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
    messages = inbox.Items #variable for items in inbox
    
    #create data files
    unread_senders_data_file_name = TEMP_DIR + "\\" + "unread_senders.txt"
    sender_data_file = open(unread_senders_data_file_name, "w+", encoding = "utf-8")

    categories_data_file_name = TEMP_DIR + "\\" + "categories.txt"
    categories_data_file = open(categories_data_file_name, "w+", encoding = "utf-8")
    
    #additional variable creation
    unread_senders_raw_list = [] #list variable to capture unread email senders with dupes
    unread_senders_unique_dict = {} #dictionary variable to capture unique unread senders with counts
    categories_senders_list = [] #dictionary variable to capture category and sender information
    flagged_messages_list = [] #list to capture flagged messege info
    
    categories_counter_int = 0
    flagged_counter_int = 0
    message_counter_int = 0

    messages.Sort("[ReceivedTime]",True)
    
    for item in all_messages_list:
        try:
            if (item['UnRead']):
                sender = item['SenderEmailAddress']
                unread_senders_raw_list.append(sender)
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
       
        # This needs some more work - commenting out for now
        # #check and store categories info
        # try:
        #     if (item['Categories']):
        #         categories_senders_list.append([item['SenderEmailAddress'],item['Categories']])
        #         categories_counter_int = categories_counter_int + 1
        # except Exception as e:
        #      append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

        #check and store flagged email info
        try:
            if (item['FlagRequest'] != ""):
                flagged_messages_dict = {} #dict to capture flagged message info

                # Assign value and add to dict
                subject = item['Subject']
                flagged_messages_dict['subject'] = subject

                sender_email = item['SenderEmailAddress']
                flagged_messages_dict['sender_email'] = sender_email

                received_time = item['ReceivedTime']
  
                flagged_messages_dict['received_time'] = received_time
                flagged_messages_list.append(flagged_messages_dict)
                flagged_counter_int = flagged_counter_int + 1
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
        
        # message_counter_int = message_counter_int + 1
        
        # if message_counter_int == 500:
        #     break
    
    unread_senders_data_gen(unread_senders_raw_list, unread_senders_unique_dict,sender_data_file)
    generate_unread_senders_viz(unread_senders_data_file_name)
    # generate_categories_viz(categories_counter_int,categories_senders_list,)
    generate_flagged_viz(flagged_counter_int, flagged_messages_list)
    word_cloud_extract(messages)
    word_cloud_display()
    
    sender_data_file.close()
    categories_data_file.close()
    


#Generates data for undread senders visualizations
def word_cloud_extract(messages):
    
    try:
        wc_file = open(WORD_CLOUD_FILE_NAME, "w+", encoding = "utf-8") #creates data file
        i = 0
        for item in messages:
            if(i<50):
                print(item.Body, file = wc_file)
                i = i + 1
            else:
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
def generate_unread_senders_viz(unread_senders_data_file_name):
    
    #Reads unread sender data and generates visualization; saves and displays
    try:
        sender_plot_file_name = TEMP_DIR + "\\" + "sender_plot.jpg"
        sender_table = pd.read_table(unread_senders_data_file_name, sep = '\t', header = None)
        plot = sender_table.groupby([0]).sum().plot(kind='pie', y=1, labeldistance=None, autopct='%1.0f%%', title="Senders of Unread Emails")
        plot.legend(bbox_to_anchor=(1,1)) #Sets legend details
        plot.set_ylabel("Senders") #Set label detail
        plot.figure.savefig(sender_plot_file_name, bbox_inches='tight') #saves plot locally
        print("\n","Top 10 Senders of Unread Emails: ", "\n", sender_table)
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
    
#Generates categories visualizations
def generate_categories_viz(categories_counter_int,categories_senders_list):
    
    try:
        #Pandas dataframe for the counted emails that are categorized
        data = {'Number of email categories': [categories_counter_int]}
        df = pd.DataFrame(data)
    
        #Removing the axis for matplotlib and creating a visual table of the counted emails that are categorize
        fig, ax = plt.subplots()
        ax.axis('off')
        ax.axis('tight')
        ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
        fig.tight_layout()
        plt.savefig("categories.jpg")
        
        #prints a tabulate table using the pandas dataframe
        print("\n")
        print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex='never'))
        print(categories_senders_list)
    except Exception as e:
        append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

#Generates flagged email visualization
def generate_flagged_viz(flagged_counter_int, flagged_messages_list):
    if (flagged_counter_int  > 0):
        try:

            # Create a dataframe from the list of dictionaries
            df = pd.DataFrame(flagged_messages_list)

        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

        try:

            df_styled = df.style.background_gradient() #adding a gradient based on values in cell
            flagged_email_list_file_name = TEMP_DIR + "\\" + "flagged_email_list.png.txt"
            # Export the data frame as an image
            dfi.export(df_styled,flagged_email_list_file_name)
            
            print("\n")
            print("Flagged Emails")
            print(tabulate(df, headers = 'keys', tablefmt = 'psql'))
            # im = Image.open("flagged_email_list.png") #displays plot in default photo viewer; can be moved to web app
            # im.show()

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
        
        #Sets hyperlink tags as indices markers
        sub1 = '<h'
        sub2 = '>'
        
        #Generates list of indices for each tag        
        indices1 = [m.start() for m in re.finditer(sub1, wc_content)]
        indices2 = [m.start() for m in re.finditer(sub2, wc_content[indices1[0]:])]
      
        
        #Prints first line of email text up to first hyperlink tag
        
        print(wc_content[0:indices1[0]],file = wc_content_cleaned)
    
        #Uses iteration through hyperlink tags to extract and print non-hyperlink text to new
        #file
        ix = 0 
        for i in range(len(indices1)-1):
            print(wc_content[indices2[ix]+1:indices1[ix+1]], file = wc_content_cleaned)
            ix = ix + 1
            
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
        data_source_str = "Outlook"

        user_input = input("Would you like to save data from Outlook to JSON file? (Y/N)")
        if user_input == "Y" or user_input == "y":
            save_data_boolean = True
        else:
            save_data_boolean = False
    elif user_input == "N" or user_input == "n":
        user_input = input("Would you like to pull data from JSON file? (Y/N)")
        if user_input == "Y" or user_input == "y":
            data_source_str = "JSON"
            save_data_boolean = False
    else:
        print("Sorry, the provided response was not understood.")

    all_messages_list = read_email(save_data_boolean,data_source_str)
    extract_email_information_from_messages_list(all_messages_list)

    #displays errors generated throughout program execution
    if ERROR_LIST != []:
        print("\n")
        print("Some errors occurred during execution:")
        for item in ERROR_LIST:
            print(item)

if __name__ == "__main__":
    main()