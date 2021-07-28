from win32com.client import Dispatch
import datetime
import speech_recognition as sr
import os
import webbrowser as wb
import random
import requests
import pandas as pd
from time import sleep
import json

def speak(str):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)

def wishMe():
    hour = int(datetime.datetime.now().hour)
    
    if hour>=0 and hour<12:
        speak("Good Morning Pratham!")
        speak("I am Jarvis Sir version 1.0 your Ai Assistant , I am Welcome you in your system!!")
        speak("Tell me What can i do for you sir?? I am Listening Your command")
        
    elif hour>=12 and hour<18:
        speak("Good Afternoon Pratham!") 
        speak("I am Jarvis Sir version 1.0 your Ai Assistant , I am Welcome you in your system!!") 
        speak("Tell me What can i do for you sir?? I am Listening Your command")
        
    else:
        speak("Good Evening Pratham!") 
        speak("I am Jarvis Sir version 1.0 your Ai Assistant , I am Welcome you in your system!!") 
        speak("Tell me What can i do for you sir?? I am Listening Your command")

time = datetime.datetime.now().strftime("%H:%M")      

def Alarm1():
    now = datetime.datetime.now().strftime("%H:%M")
    alarm = ("12:30")
    if(now == alarm):
     speak("Sir this is {} you should shut down the pc".format(now))

def Alarm2(): 
    now = datetime.datetime.now().strftime("%H:%M")
    alarm = ("18:45")
    if(now == alarm):
     speak("Sir this is {} you should shut down the pc".format(now))   

def Reminder():
    my_file = open("Reminders.txt","r")
    reminders = my_file.readlines()
    speak(reminders)  

def add_reminder():
    comands = TakeCommand()
    print(comands)
    file = open("Reminders.txt","w")
    file.write(comands)
 
def enter_event():
    with open('JARVIS Useable filse\events.json','r') as json_file:
        data = json.load(json_file)

    event_date = input("Enter Event Date: \n")
    event = input("Enter Event: \n")

    data[event_date] = [event]
    with open('JARVIS Useable filse\events.json','w') as json_file:
        json.dump(data,json_file)
    
def event_reminder():
    today = datetime.datetime.now().strftime("%d-%m")
    with open('JARVIS Useable filse\events.json','r') as json_file:
        events = json.load(json_file)
    print(events[today])
    speak(events[today])

def TakeCommand():
    r = sr.Recognizer() 
    with sr.Microphone() as source:
        print('Listening... ')
        r.pause_threshold = 1
        audio = r.listen(source)
        
        try:
           print("Recognizing...")
           task = r.recognize_google(audio,language='en-in')
           print("You said : {}".format(task))
                      
        except Exception as e:
            print("Say that again please...")
            return "None"
        return task


if __name__ == "__main__":
    # wishMe()
    # event_reminder()  

    while True:
        Alarm1()
        Alarm2()   
        task = TakeCommand()
        #here we right tasks for Jarvis!!
         
        if 'time' in task:
            strTime = datetime.datetime.now().strftime("%H:%M") 
            print(speak("sir the time is {}".format(strTime)))  

        elif 'add event' in task:
            enter_event()
            speak('Event Added sucessflly')
               
        elif 'play music' in task:
            speak("playing music sir")
            music_list = 'E:\\Pratham file\\pratham\\Song'
            songs = random.choice(os.listdir(music_list))
            play = os.path.join(music_list,songs)
            os.startfile(play)

        elif 'next' in task:
            speak("playing music sir")
            music_list = 'E:\\Pratham file\\pratham\\Song'
            songs = random.choice(os.listdir(music_list))
            play = os.path.join(music_list,songs) 
            start = os.startfile(play)

        elif 'stop' in task:
            speak("closing songs sir!")
            os.system('TASKKILL /F /IM wmplayer.exe')
        
        elif 'exit' in task:
             hour = int(datetime.datetime.now().hour)
             if hour>=0 and hour<18:
                speak("Have a good day Sir!!")
                exit()
             else:
                speak("Good Night sir!!")
                exit()

        elif 'shutdown' in task:
             hour = int(datetime.datetime.now().hour)
             if hour>=0 and hour<18:
                speak("Have a good day Sir!!")
                speak("Sir do you want to add some reminder?")
                add_reminder()
                speak("I am shuting down your pc in some secends")
                os.system("shutdown /s")
                exit()
             else:
                speak("Good Night sir!!")
                speak("Sir do you want to add some reminder?")
                add_reminder()
                speak("I am shuting down your pc in some secends")
                os.system("shutdown /s")
                exit()
                    
        elif ('open excel' == task):
            speak("opening Blank excel file sir")
            excel = "C:\\Program Files\\Microsoft Office\\Office12\\EXCEL"
            os.startfile(excel)
        
        elif 'close excel' in task:
            speak("sir first save the file otherwise data will lose")
            sleep(3)
            os.system('TASKKILL /F /IM Excel.exe')

        elif 'awesome' in task:
            speak("You are also awesome sir ")

        elif 'reminder' in task:
            Reminder()

        elif 'youtube' in task:
            speak("Opening Youtube Sir")
            wb.open_new_tab('www.youtube.com')

        elif 'facebook' in task:
            speak("Opening facebook sir!")
            wb.open('https://www.facebook.com/pratham.rathod.7127')

        elif 'insta'  in task:
            speak("Opening your instagram Account sir!")
            wb.open('https://www.instagram.com/_i_m_pratham03_/')

        elif 'channel' in task:
             speak("Opening your youtube chaneel")
             wb.open('https://www.youtube.com/channel/UCFnzIYKCQ8QGfQIBUKvbbOw?view_as=subscriber')

        elif 'jarvis' == task:
            speak("Yess sir")

        elif 'who are you' in task:
            speak("I am Jarvis Sir Your ai assistant")

        elif 'how are you jarvis' in task:
            speak("I am fine sir how are you sir")

        elif 'bandagi' in task:
            speak("saheb bandagi sir")

        elif 'open vs code' in task:
            speak("Opening Visual studio code sir")
            os.startfile('C:\\Users\\ABC\\AppData\\Local\\Programs\\Microsoft VS Code\Code.exe')
        
        elif 'close visual studio' in task:
            speak("closing Visual code sir!")
            sleep(2)
            os.system('TASKKILL /F /IM Code.exe')
                    
        elif 'open chrome' in task:
            speak("Opening chrome sir")
            wb.open("https://www.google.com")

        elif 'close google' in task:
            speak("Closing Google chrome Sir!")
            os.system('TASKKILL /F /IM chrome.exe')

        elif 'open word' == task:
            speak("opening Ms Word sir")
            os.startfile('C:\\Program Files\\Microsoft Office\\Office12\\WINWORD')

        elif 'close word' in task:
            speak("sir first save the file otherwise data will lose")
            sleep(3)
            os.system('TASKKILL /F /IM WINWORD.exe')

        elif 'play movie' in task:
            path = 'D:\\Movies'
            movie = random.choice(os.listdir(path))
            speak("Playing {} movie sir".format(movie))
            start = os.path.join(path,movie)
            os.startfile(start)

        elif 'calculator' in task:
            speak("Opening Calculater sir")
            os.system("calc")
                
        elif 'videos' in task:
            speak("Playing your edited videos sir")
            path = 'E:\Pratham file\pratham\edited videos'
            videos = random.choice(os.listdir(path))
            start = os.path.join(path,videos)
            os.startfile(start)

        elif 'close movie' in task:
            os.system('TASKKILL /F /IM vlc.exe')

        elif 'sports news' in task:
            url = "http://newsapi.org/v2/top-headlines?country=in&category=sports&apiKey=1dc79aa41b6e4f39be694888ab579dab"
            open_page = requests.get(url).json()
            articales = open_page["articles"]
            results = []
            for i in articales:
                print(results.append(i["title"]))

            for i in range(0,5):
                print(i+1,results[i])
                speak(results[i])


        elif 'live' in task:
            url = "http://newsapi.org/v2/top-headlines?country=in&apiKey=1dc79aa41b6e4f39be694888ab579dab"
            open_page = requests.get(url).json()
            articales = open_page["articles"]
            results = []
            for i in articales:
                print(results.append(i["title"]))

            for i in range(0,5):
                print(i+1,results[i])
                speak(results[i])

                   
        elif 'entertainment' in task:
            url = "https://newsapi.org/v2/top-headlines?sources=bbc-news&apiKey=1dc79aa41b6e4f39be694888ab579dab"
            open_page = requests.get(url).json()
            articales = open_page["articles"]
            results = []
            for i in articales:
                print(results.append(i["title"]))

            for i in range(0,5):
                print(i+1,results[i])
                speak(results[i])

        elif 'd drive' in task:
          change_path  = 'd://'
          speak("Opening D Drive sir")
          os.startfile(change_path)
          os.chdir(change_path) 
          
        elif 'c drive' in task:
            change_path  = 'c://'
            speak("Opening C Drive sir")
            os.startfile(change_path)
            os.chdir(change_path) 
            
        elif 'e drive' in task:
            change_path  = 'e://'
            speak("Opening E Drive sir")
            os.startfile(change_path)
            os.chdir(change_path) 

        elif 'open file' in task:
            speak("Sir which file do you want to open?")
            folder = TakeCommand() 
            try:
                os.startfile(folder)
                change_path = os.path.join(change_path,folder)
                os.chdir(change_path)
                print(os.getcwd())
                
            except Exception as e:
                print(e)
                speak("Sir thers is some probulem ,I can not find the file Please try again")

        elif 'what is today' in task:
            today = datetime.datetime.now().strftime("%d-%m")
            excel = pd.read_excel('JARVIS Useable filse\events.xlsx')
            for index,item in excel.iterrows():
                event_date = item['Date'].strftime("%d-%m")
                if event_date == today:
                    speak("sir today's reminders are as follow")
                    sleep(1)
                    speak(item['Event'])
                    print(item['Event'])