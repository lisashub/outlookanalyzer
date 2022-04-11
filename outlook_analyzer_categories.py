import win32com.client
import pandas as pd
from tabulate import tabulate
import matplotlib.pyplot as plt

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace(
    "MAPI")  # maps outlook variable to outlook application

inbox = outlook.GetDefaultFolder(6)  # outlook.GetDefaultFolder(6) is the default for the application inbox
messages = inbox.Items  # variable for items in inbox

# Checks, counts, and prints email info that have been set with categories (by user)
i = 0
for item in messages:
    if item.Categories:
        print(item.SenderEmailAddress, '-', item.Categories)
        i = i + 1
print('\n')

#Pandas dataframe for the counted emails that are categorized
data = {'Number of emails categorize': [i]}
df = pd.DataFrame(data)

#Removing the axis for matplotlib and creating a visual table of the counted emails that are categorize
fig, ax = plt.subplots()
ax.axis('off')
ax.axis('tight')
ax.table(cellText=df.values, cellLoc='center', colLabels=df.columns, loc='center')
fig.tight_layout()

#prints a tabulate table using the pandas dataframe
print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex='never'))

plt.show()