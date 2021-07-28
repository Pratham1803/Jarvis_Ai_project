import json
import datetime

today = datetime.datetime.now().strftime("%d-%m")
# today = '18-12'
with open('E:\Pratham file\Coding programms\Python\Python Projects\J.A.R.V.I.S\JARVIS Useable filse\events.json','r') as json_file:
    data = json.load(json_file)
    # data = data['Birthdays']

# print(data[today])

event_date = input("Enter Event Date: \n")
event = input("Enter Event: \n")

data[event_date] = [event]
with open('E:\Pratham file\Coding programms\Python\Python Projects\J.A.R.V.I.S\JARVIS Useable filse\events.json','w') as json_file:
    json.dump(data,json_file)

print(data)