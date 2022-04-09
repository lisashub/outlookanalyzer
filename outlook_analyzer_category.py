import win32com.client
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from tabulate import tabulate
from PIL import Image

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace(
    "MAPI")  # maps outlook variable to outlook application

inbox = outlook.GetDefaultFolder(6)  # outlook.GetDefaultFolder(6) is the default for the application inbox
messages = inbox.Items  # variable for items in inbox

i = 0  # count for emails that are categorize
category_list = []
for item in messages:
    try:
        if item.Categories:
            i = i + 1  # count for total categories

            # Adds Sender Email Address into dictionary
            category_dict = {}
            sender = item.SenderEmailAddress
            category_dict['Sender_Email_Address'] = sender
            category_list.append(category_dict)

            # Adds category name to category list
            category = item.Categories
            category_list.append(category)
    except Exception:
        print("Error extracting categorize emails")

# Count for emails specific color category
b = 0
g = 0
o = 0
p = 0
r = 0
y = 0
for category in category_list:
    try:
        if category == 'Blue category':
            b = b + 1
        if category == 'Green category':
            g = g + 1
        if category == 'Orange category':
            o = o + 1
        if category == 'Purple category':
            p = p + 1
        if category == 'Red category':
            r = r + 1
        if category == 'Yellow category':
            y = y + 1
    except Exception:
        print('Error with counting categorize emails')

print(category_list)  # prints category list
print('Number of email(s) categorize:', i)  # print total number or emails categorize

# Pandas dataframe for the counted emails that are categorized
data = {'Blue category': [b], 'Green category': [g], 'Orange category': [o], 'Purple category': [p],
        'Red category': [r],
        'Yellow category': [y]}
df = pd.DataFrame(data)

# Removing the axis and creating a visual table of the counted emails that are categorize
fig, ax = plt.subplots()
plt.title('Number of Email(s) Categorize')
ax.axis('off')
ax.axis('tight')
table = ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
table.scale(1, 2)

plt.savefig('Email_Category.png', facecolor=fig.set_facecolor('#f8f8ff'))  # saves plot locally
im = Image.open('Email_Category.png')  # displays plot in default photo viewer; can be moved to web app
# im.show()

# resize_im = im.resize((500,500)) #line 78 & 79 was testing on resizing image to fit PDF
# resize_im.save('category_resize.png')

# prints a tabulate table using the pandas dataframe
print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex='never'))

# Testing code for creating PDF analytic dashboard
pdf = FPDF()
pdf.add_page()  # adds pdf page
pdf.set_font('Arial', 'B', 16)  # sets pdf fonts
pdf.cell(0, 60, 'Outlook Analzyer Dashboard', 0, 0, align='C')  # creates a cell (retangle area) w/ text
pdf.image('Email_Category.png', x=0, y=80, w=210, h=150)  # puts and positions an image into the pdf
pdf.output('outlook_analyzer.pdf', 'F')  # saves pdf into local file
pdf.open()
