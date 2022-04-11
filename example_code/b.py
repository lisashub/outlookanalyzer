end_date = date.today() - relativedelta(months=+1)
start_date = date.today() - relativedelta(months=+2)

start_date_str = start_date.strftime('%m/%d/%Y %H:%M %p')
end_date_str = end_date.strftime('%m/%d/%Y %H:%M %p')

print("Staring date: " + str(start_date))
print("Ending date: " + str(end_date))

filtered_message = messages.Restrict("[ReceivedTime] >= '" + end_date_str + "' AND [ReceivedTime] <= '" + start_date_str + "'")
