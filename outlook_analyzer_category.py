import win32com.client
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from tabulate import tabulate
from PIL import Image
from tqdm import tqdm

def unique(list1): #custom function to identify unique senders
    unique_list = []
    for item in list1:
        if item not in unique_list:
            unique_list.append(item)
    return unique_list

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  # maps outlook variable to outlook application
inbox = outlook.GetDefaultFolder(6)  # outlook.GetDefaultFolder(6) is the default for the application inbox
messages = inbox.Items  # variable for items in inbox
categories_data_file = open("categories_data.txt", "w+", encoding = "utf-8") #creates data file

categories_counter_int = 0
category_list = []
category_dict = {}

#Streamlined process of collecting categories
for item in tqdm(messages):
    try:
        if item.Categories:
            item_categories = item.Categories.split(",")
            for category in item_categories:
                category_list.append(category)
    except Exception:
        print("Error extracting categorize emails")

unique_categories = unique(category_list) #sends category list to function named "unique" and saves list of unique values to variable
for category in unique_categories: #loops through unique categories and counts occurrances; saves results into category_dict
    category_dict[category] = category_list.count(category)

print("\n")
print('Number of categories:', len(unique_categories))  # print total number or emails categorize

#NEW: Stores data into text file
for item in category_dict.items():
        print(item[0], "\t", item[1], file = categories_data_file)
        categories_counter_int = categories_counter_int + 1
categories_data_file.close()

# Pandas dataframe for the counted emails that are categorized
df = pd.read_csv("categories_data.txt", sep = "\t")
# Removing the axis and creating a visual table of the counted emails that are categorize
fig, ax = plt.subplots()
plt.title('Number of Email(s) Categories')
ax.axis('off')
ax.axis('tight')
table = ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
table.scale(1, 2)
plt.savefig('Email_Category.png')  # saves plot locally, facecolor=fig.set_facecolor('#f8f8ff') - changes background color
im = Image.open('Email_Category.png')  # displays plot in default photo viewer; can be moved to web app
# im.show()
# resize_im = im.resize((500,500)) #line 78 & 79 was testing on resizing image to fit PDF
# resize_im.save('category_resize.png')

# prints a tabulate table using the pandas dataframe
print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex='never'))

