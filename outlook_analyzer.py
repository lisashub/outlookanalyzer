
import win32com.client #core extraction library
from tqdm import tqdm # library to display extraction progress bar
import pandas as pd # library to tabulate data and generate plot/image
from tabulate import tabulate
import matplotlib.pyplot as plt
import dataframe_image as dfi
from wordcloud import WordCloud, STOPWORDS #word cloud generation library
import re #regex library used to clean data




def main():
    
    #Connection to Outlook object model established; can move to extract only module, too
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
    inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
    messages = inbox.Items #variable for items in inbox
    
    ERROR_LIST = []
    
    
    print("\n")
    print("Welcome to Outlook Analyzer!")
    user_response = input("Would you like to extract fresh data from Outlook? (Y/N)")
    
    #Logic to determine whether extraction is required
    if user_response == "Y" or user_response == "y":
        ERROR_LIST = extract_outlook_information(messages, ERROR_LIST)
    elif user_response == "N" or user_response == "n":
        #placeholder code; to re-create to hook in PDF generation logic using existing files
        try:
            word_cloud_extract(messages, ERROR_LIST)
            word_cloud_display(ERROR_LIST)
        except Exception as e:
            ERROR_LIST.append("error extracting generating visualizations: " + str(e))
            
    else:
        print("Sorry, the provided response was not understood.")
    
    #displays errors generated throughout program execution
    if ERROR_LIST != []:
        print("\n")
        print("Some errors occurred during execution:")
        for item in ERROR_LIST:
            print(item)

def extract_outlook_information(messages, ERROR_LIST): #to modify as new features required
    
    #create data files
    sender_data_file = open("unread_senders.txt", "w+", encoding = "utf-8")
    categories_data_file = open("categories.txt", "w+", encoding = "utf-8")
    
    
    #additional variable creation
    unread_senders_raw_list = [] #list variable to capture unread email senders with dupes
    unread_senders_unique_dict = {} #dictionary variable to capture unique unread senders with counts
    categories_senders_list = [] #dictionary variable to capture category and sender information
    flagged_messages_list = [] #list to capture flagged messege info
    flagged_messages_dict = {} #dict to capture flagged message info
    
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
            ERROR_LIST.append("error extracting details for unread email senders:" + str(e))
        
        #check and store categories info
        try:
            if item.Categories: #checks and stores info for emails that have been set with categories (by user)
                categories_senders_list.append([item.SenderEmailAddress,item.Categories])
                categories_counter_int = categories_counter_int + 1
        except Exception as e:
            ERROR_LIST.append("error extracting details for categories:" + str(e))
        
        #check and store flagged email info
        try:
            if (item.FlagRequest != ""): #checks and stores flag info
             
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
            ERROR_LIST.append("error extracting details for flagged email:" + str(e))
        
        message_counter_int = message_counter_int + 1
        
        if message_counter_int == 500:
            break
    
    ERROR_LIST = unread_senders_data_gen(unread_senders_raw_list,ERROR_LIST, unread_senders_unique_dict,sender_data_file)
    ERROR_LIST = generate_unread_senders_viz(ERROR_LIST)
    ERROR_LIST = generate_categories_viz(categories_counter_int,categories_senders_list, ERROR_LIST)
    ERROR_LIST = generate_flagged_viz(flagged_counter_int, flagged_messages_list, ERROR_LIST)
    ERROR_LIST = word_cloud_extract(messages, ERROR_LIST)
    ERROR_LIST = word_cloud_display(ERROR_LIST)
    
    sender_data_file.close()
    categories_data_file.close()
    
    return (ERROR_LIST)

#Generates data for undread senders visualizations
def word_cloud_extract(messages, ERROR_LIST):
    messages.Sort("[ReceivedTime]",True)
    wc_file = open("word_cloud_text.txt", "w+", encoding = "utf-8") #creates data file
    i = 0
    for item in messages:
        if(i<50):
            print(item.Body, file = wc_file)
            i = i + 1
        else:
            word_cloud_content_clean(ERROR_LIST) #text-cleaning function called
            wc_file.close()
            return (ERROR_LIST)
        
#Generates data for unread senders visualizations
def unread_senders_data_gen(unread_senders_raw_list,ERROR_LIST, unread_senders_unique_dict,sender_data_file):
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
    
    return (ERROR_LIST)

#Generates unread senders visualizations
def generate_unread_senders_viz(ERROR_LIST):
    
    #Reads unread sender data and generates visualization; saves and displays
    sender_table = pd.read_table('unread_senders.txt', sep = '\t', header = None)
    plot = sender_table.groupby([0]).sum().plot(kind='pie', y=1, labeldistance=None, autopct='%1.0f%%', title="Senders of Unread Emails")
    plot.legend(bbox_to_anchor=(1,1)) #Sets legend details
    plot.set_ylabel("Senders") #Set label detail
    plot.figure.savefig("sender_plot.jpg", bbox_inches='tight') #saves plot locally
    print("\n","Top 10 Senders of Unread Emails: ", "\n", sender_table)
    
    return(ERROR_LIST)

#Generates categories visualizations
def generate_categories_viz(categories_counter_int,categories_senders_list, ERROR_LIST):
    
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
    
    return(ERROR_LIST)

#Generates flagged email visualization
def generate_flagged_viz(flagged_counter_int, flagged_messages_list, ERROR_LIST):
    if (flagged_counter_int  > 0):
        try:

            # Create a dataframe from the list of dictionaries
            df = pd.DataFrame(flagged_messages_list)

        except Exception as e:
            ERROR_LIST.append("error when creating data frame:" + str(e))

        try:

            df_styled = df.style.background_gradient() #adding a gradient based on values in cell
            # Export the data frame as an image
            dfi.export(df_styled,"flagged_email_list.png")
            
            print("\n")
            print("Flagged Emails")
            print(tabulate(df, headers = 'keys', tablefmt = 'psql'))
            # im = Image.open("flagged_email_list.png") #displays plot in default photo viewer; can be moved to web app
            # im.show()

        except Exception as e:
            ERROR_LIST.append("error when exporting data frame:" + str(e))
            
            #Prints some ouputs to Command Line
    print("\n")
    print("Number of Flagged emails: ", flagged_counter_int)
    
    return(ERROR_LIST)


#Removes hyperlink information from email body to produce more meaningful clouds
def word_cloud_content_clean(ERROR_LIST):
    wc_content= open("word_cloud_text.txt", "r", encoding = "utf-8").read()
    wc_content_cleaned = open("word_cloud_text_cleaned.txt", "w+", encoding = "utf-8")
    
    #Sets hyperlink tags as indices markers
    sub1 = '<'
    sub2 = '>'
    
    #Generates list of indices for each tag        
    indices1 = [m.start() for m in re.finditer(sub1, wc_content)]
    indices2 = [m.start() for m in re.finditer(sub2, wc_content)]
    
    #Prints first line of email text up to first hyperlink tag
    print(wc_content[0:indices1[0]],file = wc_content_cleaned)

    #Uses iteration through hyperlink tags to extract and print non-hyperlink text to new
    #file
    ix = 0 
    for i in range(len(indices1)-1):
        print(wc_content[indices2[ix]+1:indices1[ix+1]], file = wc_content_cleaned)
        ix = ix + 1
        
    wc_content_cleaned.close()
    return(ERROR_LIST)

#Generates word cloud from cleaned email text
def word_cloud_generate():
    wc_content= open("word_cloud_text_cleaned.txt", "r", encoding = "utf-8").read()
    stop_words = ["said", "email", "s", "will", "u", "re"] + list(STOPWORDS) #customized stopword list
    wordcloud = WordCloud(stopwords = stop_words).generate(str(wc_content))
    return(wordcloud)

#Saves and displays newly generated word cloud to user
def word_cloud_display(ERROR_LIST):
    try:
        plt.clf()
        plt.imshow(word_cloud_generate())
        plt.axis('off')
        plt.savefig('word_cloud.jpg')
    except Exception as e:
        ERROR_LIST.append("error generating word cloud display: " + str(e))
    return(ERROR_LIST)

#Identifiies unique elements within a list
def unique (list1):
    unique_elements_list = []
    for item in list1:
        if item not in unique_elements_list:
            unique_elements_list.append(item)
    return unique_elements_list

if __name__ == "__main__":
    main()