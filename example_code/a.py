e_dt = date.today() - relativedelta(months=+1)
s_dt = date.today() - relativedelta(months=+2)

start_dt = s_dt.strftime('%m/%d/%Y %H:%M %p')
end_dt = e_dt.strftime('%m/%d/%Y %H:%M %p')

print("Staring date: " + str(s_dt))
print("Ending date: " + str(e_dt))

filtered_message = messages.Restrict("[ReceivedTime] >= '" + start_dt + "' AND [ReceivedTime] <= '" + end_dt + "'")
