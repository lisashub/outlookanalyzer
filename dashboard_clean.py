import os
from fpdf import FPDF
from datetime import datetime

# sets real time date as a string
now = datetime.now()
date = now.strftime('%m/%d/%y')

# Code for creating PDF dashboard
pdf = FPDF()
pdf.add_page()  # adds pdf page
pdf.set_font('Arial', 'B', 18)  # sets pdf fonts
pdf.cell(0, 60, 'Outlook Analzyer Dashboard', 0, 0, align='C')  # Puts in title
pdf.cell(-190, 75, date, 0, 0, align='C') # Puts in real time date
pdf.image('icon.png', x=0, y=0, w=35, h=30) # Puts in icon
pdf.image('true_blue.jpg', x=35, y=0, w=175, h=30) # Puts in and positions header
pdf.image('sender_plot.jpg', x=0, y=75, w=210) # Puts in and positions send plot data
pdf.image('Email_Category.png', x=0, y=130, w=210/2)  # Puts in and positions category
pdf.image('flagged_email_list.png', x=105, y=130, w=210/2) # Puts in and positions flagged email data
pdf.image('word_cloud.jpg', x=0, y=200, w=210, h=90) # Puts in and positions word cloud
pdf.output('dashboard.pdf', 'F')  # saves pdf into local file
pdf.open()


# Code for cleaning directory
dir_name = '' # User's directory path
dir_list = os.listdir(dir_name) #returns the list of all files and directories in the specified path

# Goes through directory and removes/cleans the files with specified extensions (i.e .txt, .tmp, .png, etc.)
while True:
        dir_option = input('Would you like to clean your directory? (y/n):')
        if dir_option == 'y' or dir_option == 'Y':
            for item in dir_list:
                if item.endswith('.txt') # or item.endswith('.png') or item.endswith('.jpg'): comment this out just in case you don't want to remove the images in the PDF code above
                    os.remove(os.path.join(dir_name, item))
            print('Directory cleaned')
            break;
        elif dir_option == 'n' or dir_option == 'N':
            pass
            break;


