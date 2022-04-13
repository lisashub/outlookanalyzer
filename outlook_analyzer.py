
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
FLAGGED_EMAIL_LIST_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "flagged_email_list.png"
SENDER_PLOT_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "sender_plot.jpg"
CATEGORIES_IMAGE_FILE_NAME = TEMP_DIR + "\\" + TIME_STR + "_" + "categories.jpg"

def append_to_error_list(function_name, error_text):
    ERROR_LIST.append("function: " + function_name + " | " +  "error: " + error_text)

#Extracts data from Outlook
def extract_outlook_information(): #to modify as new features required

    #Connection to Outlook object model established
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
    inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
    messages = inbox.Items #variable for items in inbox

    #create data files

    sender_data_file = open(UNREAD_SENDERS_DATA_FILE_NAME, "w+", encoding = "utf-8")
    categories_data_file = open(CATEGORIES_DATA_FILE_NAME, "w+", encoding = "utf-8")
    
    #additional variable creation
    unread_senders_raw_list = [] #list variable to capture unread email senders with dupes
    unread_senders_unique_dict = {} #dictionary variable to capture unique unread senders with counts
    categories_senders_list = [] #dictionary variable to capture category and sender information
    flagged_messages_list = [] #list to capture flagged messege info
    
    categories_counter_int = 0
    flagged_counter_int = 0
    message_counter_int = 0

    messages.Sort("[ReceivedTime]",True)
    
    for item in tqdm(messages):

                #check and store unread email info
        try:
            if (item.Unread == True):
                sender = item.SenderEmailAddress
                unread_senders_raw_list.append(sender)
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
        
       #check and store categories info
        try:
            if item.Categories: #checks and stores info for emails that have been set with categories (by user)
                categories_senders_list.append([item.SenderEmailAddress,item.Categories])
                categories_counter_int = categories_counter_int + 1
        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

        #check and store flagged email info
        try:
            if (item.FlagRequest != ""): #checks and stores flag info - TO DO: Fix this to not get resolved flags
            
             flagged_messages_dict = {} #dict to capture flagged message info
            
             # Assign value and add to dict
             subject = item.subject
             flagged_messages_dict['subject'] = subject

             sender_email = item.SenderEmailAddress
             flagged_messages_dict['sender_email'] = sender_email

             received_time = item.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S")
             flagged_messages_dict['received_time'] = received_time

             flagged_messages_list.append(flagged_messages_dict)
             
             flagged_counter_int = flagged_counter_int + 1

        except Exception as e:
            append_to_error_list(str(sys._getframe().f_code.co_name),str(e))
        
        message_counter_int = message_counter_int + 1
        
        if message_counter_int == 500:
            break
    
    unread_senders_data_gen(unread_senders_raw_list, unread_senders_unique_dict,sender_data_file)
    generate_unread_senders_viz()
    generate_categories_viz(categories_counter_int,categories_senders_list)
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
        plt.savefig(CATEGORIES_IMAGE_FILE_NAME)
        
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
            # Export the data frame as an image
            dfi.export(df_styled,FLAGGED_EMAIL_LIST_FILE_NAME)
            
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
        # TO DO: Add more input questions for time range and max number of emails to pull

        extract_outlook_information()

    # # Commented out for now - To Lisa: What is the purpose of this elif block? We need to connect to Outlook to get the data for the world cloud right?
    # elif user_input == "N" or user_input == "n":
    #     #placeholder code; to re-create to hook in PDF generation logic using existing files
    #     try:
    #         word_cloud_display()
    #     except Exception as e:
    #         append_to_error_list(str(sys._getframe().f_code.co_name),str(e))

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