import json

TEMP_DIR = "C:\WINDOWS\Temp"
FILE_NAME = "email_out"
FILE_PATH = TEMP_DIR + "\\" + FILE_NAME

# with open('data.json') as d:
with open(FILE_PATH) as d:
    all_messages_list = json.load(d)
    # print(all_messages_list)
    # count_int = (len(all_messages_list))


print("\nPrinting nested dictionary as a key-value pair\n")
for i in all_messages_list:

    # print(i)
    print("SenderEmailAddress:", i['SenderEmailAddress'])



    # print(type(dictData))
    # print(dictData['games'])