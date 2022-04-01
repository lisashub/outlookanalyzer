#Extracts body text from 50 most recent emails and calls text-cleaning function
def word_cloud_extract(messages):
    print("Extracting Word Cloud Content...")
    messages.Sort("[ReceivedTime]",True)
    wc_file = open("word_cloud_text.txt", "w+", encoding = "utf-8") #creates data file
    i = 0
    for item in messages:
        if(i<50):
            print(item.Body, file = wc_file)
            i = i + 1
        else:
            word_cloud_content_clean() #text-cleaning function called
            wc_file.close()
            return

#Removes hyperlink information from email body to produce more meaningful clouds
def word_cloud_content_clean():
    wc_content= open("word_cloud_text.txt", "r", encoding = "utf-8").read()
    wc_content_cleaned = open("word_cloud_text_cleaned.txt", "w+", encoding = "utf-8")
    
    #Sets hyperlink tags as indices markers
    sub1 = '<'
    sub2 = '>'
    
    #Generates list of indices for each tag        
    indices1 = [m.start() for m in re.finditer(sub1, wc_content)]
    indices2 = [m.start() for m in re.finditer(sub2, wc_content)]
    print(len(indices1))
    print(len(indices2))
    
    #Prints first line of email text up to first hyperlink tag
    print(wc_content[0:indices1[0]],file = wc_content_cleaned)

    #Uses iteration through hyperlink tags to extract and print non-hyperlink text to new
    #file
    ix = 0 
    for i in range(len(indices1)-1):
        print(wc_content[indices2[ix]+1:indices1[ix+1]], file = wc_content_cleaned)
        ix = ix + 1
        
    wc_content_cleaned.close()

#Generates word cloud from cleaned email text
def word_cloud_generate():
    print("Generating Word Cloud...")
    wc_content= open("word_cloud_text_cleaned.txt", "r", encoding = "utf-8").read()
    stop_words = ["said", "email", "s", "will", "u", "re"] + list(STOPWORDS) #customized stopword list
    wordcloud = WordCloud(stopwords = stop_words).generate(str(wc_content))
    return(wordcloud)

#Displays word cloud to user
def word_cloud_display():
    print("Displaying Word Cloud...")
    plt.imshow(word_cloud_generate())
    plt.axis('off')
    plt.savefig('word_cloud.jpg')

    

import win32com.client #core extraction library
from wordcloud import WordCloud, STOPWORDS #word cloud generation library
import matplotlib.pyplot as plt #word cloud display library
import re #regex library used to clean data


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
messages = inbox.Items #variable for items in inbox


word_cloud_extract(messages)
word_cloud_display()



