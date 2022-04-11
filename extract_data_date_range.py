import win32com.client #core extraction library
from tqdm import tqdm # library to display extraction progress bar
import json
from datetime import timedelta
from datetime import date
from dateutil.relativedelta import relativedelta

TEMP_DIR = "C:\WINDOWS\Temp"
FILE_NAME = "email_out"
FILE_PATH = TEMP_DIR + "\\" + FILE_NAME

def read_email(save_data_boolean,data_source_str):

    if data_source_str == "Outlook":

        print("Connecting to Outlook...")

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #maps outlook variable to outlook application
        inbox = outlook.GetDefaultFolder(6) #outlook.GetDefaultFolder(6) is the default for the application inbox
        messages = inbox.Items #variable for items in inbox

        # Example using timedelta by days
        received_dt = date.today() - timedelta(days=30)
        # e_dt = date.today() - timedelta(days=30)
        # s_dt = date.today() - timedelta(days=60)

        # Example using relativedelta by month
        end_date = date.today() - relativedelta(months=+1)
        start_date = date.today() - relativedelta(months=+2)

        start_date_str = start_date.strftime('%m/%d/%Y %H:%M %p')
        end_date_str = end_date.strftime('%m/%d/%Y %H:%M %p')

        print("Extracting email in this date range: ")
        print("Staring date: " + str(start_date))
        print("Ending date: " + str(end_date))

        # Other ways to restrict email by ReceivedTime
        # received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
        # filtered_message = messages.Restrict("[ReceivedTime] >= '" + str(received_dt) + "'")

        filtered_message = messages.Restrict("[ReceivedTime] >= '" + start_date_str + "' AND [ReceivedTime] <= '" + end_date_str + "'")

        # today = datetime.date.today()
        # first = today.replace(day=1)
        # lastMonth = first - datetime.timedelta(days=1)
        # print(lastMonth.strftime("%Y%m"))

        # Can also restrict emails in other ways
        # current_message = messages.Restrict("[SenderEmailAddress] = 'notifications@github.com' ")

        # Get the most recent email
        # current_message = messages.GetLast()
        # print(current_message)

        all_messages_list = []

        messages.Sort("[ReceivedTime]",True)
        for item in tqdm(filtered_message):
            messages_dict = {}

            # Some attributes are only applicable to email, not other items such as meeting invites
            if item.Class == 43: # normal email are class 43 (meeting invites are 53, responses are 56 )

                try:
                    alternate_recipient_allowed = item.AlternateRecipientAllowed
                    messages_dict['AlternateRecipientAllowed'] = alternate_recipient_allowed
                except Exception as e:
                    print("error extracting:" + str(e))

                try:
                    is_marked_as_task = item.IsMarkedAsTask
                    messages_dict['IsMarkedAsTask'] = is_marked_as_task
                except Exception as e:
                    print("error extracting:" + str(e))

                try:
                    read_receipt_requested = item.ReadReceiptRequested
                    messages_dict['ReadReceiptRequested'] = read_receipt_requested
                except Exception as e:
                    print("error extracting:" + str(e))

                try:
                    to = item.To
                    messages_dict['To'] = to
                except Exception as e:
                    print("error extracting:" + str(e))

                try:
                    cc = item.CC
                    messages_dict['CC'] = cc
                except Exception as e:
                    print("error extracting:" + str(e))

                try:
                    bcc = item.BCC
                    messages_dict['BCC'] = bcc
                except Exception as e:
                    print("error extracting:" + str(e))

            # Continue with other attributes applicable to email and other items such as meeting invites

            try:
                attachments_list = []
                # Loops over each attachment and adds a list of attachment file names
                for attachment in item.Attachments:
                    attachments_list.append(attachment.FileName)
                messages_dict['Attachments'] = attachments_list
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                auto_forwarded = item.AutoForwarded
                messages_dict['AutoForwarded'] = auto_forwarded
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                email_class = item.Class
                messages_dict['Class'] = email_class
            except Exception as e:
                print("error extracting:" + str(e))

        
            try:
                # Split the CSV string from Categories into a list 
                messages_dict['Categories'] = item.Categories.split(",")
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                companies = item.Companies
                messages_dict['Companies'] = companies
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                importance = item.Importance
                messages_dict['Importance'] = importance
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                sender_name = item.SenderName
                messages_dict['SenderName'] = sender_name
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                sender_email_address = item.SenderEmailAddress
                messages_dict['SenderEmailAddress'] = sender_email_address
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                sent_on = item.SentOn
                messages_dict['SentOn'] = sent_on
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                sender_email_type = item.SenderEmailType
                messages_dict['SenderEmailType'] = sender_email_type
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                reminder_set = item.ReminderSet
                messages_dict['ReminderSet'] = reminder_set
            except Exception as e:
                print("error extracting:" + str(e))

            
            try:
                recipients_list = []
                # Loops over each recipient and adds a list of recipient  names
                for recipient in item.Recipients:
                    recipients_list.append(recipient.Name)
                messages_dict['Recipients'] = recipients_list
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                sensitivity = item.Sensitivity
                messages_dict['Sensitivity'] = sensitivity
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                flag_request = item.FlagRequest
                messages_dict['FlagRequest'] = flag_request
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                unread = item.UnRead
                messages_dict['UnRead'] = unread
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                size = item.Size
                messages_dict['Size'] = size
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                subject = item.Subject
                messages_dict['Subject'] = subject
            except Exception as e:
                print("error extracting:" + str(e))
            
            try:
                body = item.Body
                messages_dict['Body'] = body
            except Exception as e:
                print("error extracting:" + str(e))

            try:
                received_time = item.ReceivedTime
                # received_time = item.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S")
                messages_dict['received_time'] = received_time
            except Exception as e:
                print("error extracting:" + str(e))

            all_messages_list.append(messages_dict)

        # Save email data to disk as a list of JSON key value items
        if save_data_boolean:
            print("Saving Outlook email data to: " + FILE_PATH)
            json_object = json.dumps(all_messages_list, indent = 4, sort_keys=True, default=str)

            with open(FILE_PATH, "w+") as outfile:
                outfile.write(json_object)
                    
    elif data_source_str == "JSON":

        # Opening JSON file
        with open(FILE_PATH) as json_file:
            all_messages_list = json.load(json_file)
        
    # Here the output should be the same if it comes from Outlook or comes from the JSON file
    print("\nPrinting values from all_messages_list\n")
    for dict_item in all_messages_list:
        # Example of extracting something from the list of dict_item
        print("SenderEmailAddress:", dict_item['SenderEmailAddress'])

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

    read_email(save_data_boolean,data_source_str)

if __name__ == "__main__":
    main()