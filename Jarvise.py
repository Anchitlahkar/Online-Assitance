import cgitb
import cv2
import datetime
import os
# import Tools.scripts.google
import smtplib
import webbrowser
from time import ctime

import pyttsx3
import speech_recognition as sr
import wikipedia
import win32com.servers.dictionary
from googlesearch import search
import math
from email.mime import audio

cgitb.enable()


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[0].id)
engine.setProperty('voice', voices[0].id)



def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now(). hour)
    if hour >= 0 and hour < 12:
        speak("Good morning")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon")

    else:
        speak("Good Evening")

    speak('my name is Jarvice. how can i help you')

def takeCommand(ask=False):
    # It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone()as source:
        speak("Hearing Sir")
        print("Hearing Sir...")
        r.pause_threshold = 1
        if ask:
            print(ask)
        audio = r.listen(source)

    try:
        print("recognizing...")
        quary = r.recognize_google(audio, language='hi=in')
        print(f"User said: {quary}\n") 

    except Exception as e:
        # print(e)
        speak("Sorry Sir can you repeat that one")
        print("Say that again Please...")
        return "None"
    return quary


if __name__ == '__main__':
    wishMe()
    while True:
        query = takeCommand().lower()

        # logic for executing task based on query
        if 'wikipedia' in query:
            speak("Searching Wikipedia....")
            query = query.replace("wikipedia", "")
            url = 'https://www.wikipedia.org/'
            results = wikipedia.summary(query, sentences=5)
            speak("according to wikipedia")
            webbrowser.get().open(url)
            print(results)
            speak(results)
            engine.runAndWait()
            engine.runAndWait()
            engine.runAndWait()


        
        # elif 'open youtube' in query:
        #     speak("opening Youtube")
        #     webbrowser.open("youtube.com")

        elif 'open youtube' in query:
            speak('Opening youtube')
            url = 'https://www.youtube.com/'
            webbrowser.get().open(url)
            engine.runAndWait()
            engine.runAndWait()


        # elif 'open google' in query:
        #     speak("opening google")
        #     webbrowser.open("google.com")

        elif 'open google' in query:
            speak('Opening google')
            url = 'https://www.google.com/'
            webbrowser.get().open(url)
            engine.runAndWait()
            engine.runAndWait()



        # elif 'open nasa' in query:
        #     speak("opening Nasa.gov")
        #     webbrowser.open("nasa.gov")

        elif 'open nasa' in query:
            speak('Opening nasa')
            url = 'https://www.nasa.gov/'
            webbrowser.get().open(url)
            engine.runAndWait()
            engine.runAndWait()
            

    #    elif 'play music' in query:
    #        music_dic = "D:\\ wikimusic"
    #        songs = os.listdir(music_dic)
    #        print(songs)
    #        os.startfile(os.path.join(music_dic, songs[0]))

        elif 'play music' in query:
            speak('playing music')
            url = 'https://wynk.in/music'
            webbrowser.get().open(url)
            engine.runAndWait()
            engine.runAndWait()

            

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, The time is {strTime}")
            engine.runAndWait()


        elif 'open microsoft team' in query :
            speak('opening Microsoft team')
            codePath = "C:\\Users\\lenevo\\AppData\\Local\\Microsoft\\Teams"
            os.startfile(codePath)
            engine.runAndWait()
            engine.runAndWait()


        elif "open WhatApp" in query :
            speak("opening whatapp")
            codePath = "C:\\Users\\lenevo\\AppData\\Local\\WhatsApp\\WhatsApp.exe"
            os.startfile(codePath)
            engine.runAndWait()
            engine.runAndWait()

        
        elif "open zoom" in query :
            speak("opening whatapp")
            codePath = "C:\\Users\\lenevo\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe"
            os.startfile(codePath)
            engine.runAndWait()
            engine.runAndWait()



        elif "open rammarly" in query :
            speak("opening Grammarly")
            codePath = "C:\\Users\\lenevo\\AppData\\Local\\GrammarlyForWindows\\GrammarlyForWindows.exe"
            os.startfile(codePath)
            engine.runAndWait()
            engine.runAndWait()


        elif "open office" in query :
            speak("opening Office")
            codePath = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome_proxy.exe"
            os.startfile(codePath)
            engine.runAndWait()
            engine.runAndWait()



        elif 'open programming' in query:
            speak("opening visiul studio code")
            codePath = "C:\\Users\\lenevo\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
            engine.runAndWait()
            engine.runAndWait()


        elif 'open microsoft' in query:
            speak('opening Microsoft Edge')
            searchPath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
            os.startfile(searchPath)
            engine.runAndWait()
            engine.runAndWait()


        elif 'open unity' in query:
            speak("opening unity")
            codePath = "F:\\Unitygaming\\2019.4.2f1\\Editor\\Unity.exe"
            os.startfile(codePath)
            engine.runAndWait()
            engine.runAndWait()


        # elif 'play music' in query:
        #     speak("playing Music")
        #     chromePath = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome_proxy.exe  --profile-directory=Profile 1 --app-id=emacdpakoihlgkpbekbfnhinbipjbepd"
        #     os.startfile(chromePath)


        elif 'what is your name' in query:
            speak('Sir My name is Jarvice.. And I am an Artificial Itelligents, I only work for my Sir Anchit Lahkar.')
            engine.runAndWait()

        elif 'who are you' in query:
            speak('Sir My name is Jarvice.. And I am an Artificial Itelligents I only work for my Sir Anchit Lahkar.')
            engine.runAndWait()

        elif 'what is python' in query:
            speak('sir please import codes')
            engine.runAndWait()

        elif 'the date' in query:
            print(ctime())
            speak(ctime())
            engine.runAndWait()

        elif 'what' in query:
            url = 'http://google.com/search?q=' + query
            webbrowser.get().open(url)
            speak('hear is what i found for'+ query)
            speak(webbrowser.get().open(url))
            engine.runAndWait()
            engine.runAndWait()
            engine.runAndWait()
         


        elif 'where' in query:
            url = 'http://google.nl/maps/place/' + query + '/&amp;'
            webbrowser.get().open(url)
            print('Here is the location for ' + query)
            speak('Here is the location for ' + query)
            print("\n")
            engine.runAndWait()
            engine.runAndWait()
            engine.runAndWait()



        elif 'answer' in query:
            speak('what do you want to search for???')
            search = takeCommand('what do you want to search for???')
            speak('Searching Wikipedia')
            url = 'https://en.wikipedia.org/wiki/' + search
            webbrowser.get().open(url)
            speak('Here is what I found for' + search)
            print('Here is what I found for' + search)
            engine.runAndWait()
            engine.runAndWait()
            engine.runAndWait()


        
        elif 'search microsoft' in query:
            speak('what do you want to search for???')
            search = takeCommand('what do you want to search for???')
            url = 'https://www.bing.com/search?q=' + search
            webbrowser.get().open(url)
            speak('Here is what I found for' + search)
            print('Here is what I found for' + search)
            engine.runAndWait()
            engine.runAndWait()
            engine.runAndWait()


        elif 'open spotify' in query:
            speak('opening spotify.com')
            url = 'https://open.spotify.com/'
            webbrowser.get().open(url)
            speak('now you can hear music on spotify.')
            engine.runAndWait()
            engine.runAndWait()
            engine.runAndWait()



        elif 'open online website' in query:
            speak('opening wix.com')
            url = 'https://www.wix.com/account/sites?utm_source=email_mkt&utm_campaign=em_welcome-to-wix_en&experiment_id=button_cta_1_desktop'
            webbrowser.get().open(url)
            speak('Now you can code websites')
            engine.runAndWait()
            engine.runAndWait()
            engine.runAndWait()

        elif 'talk to me' in query :
            speak('activating selftalk')
            speak("how can i help you through google as search engine")
            search = takeCommand("how can i help you through google as search engine")
            url = 'http://google.com/search?q=' + search
            webbrowser.get().open(url)
            speak("hear's what i found")
            speak(webbrowser.open(url))
            engine.runAndWait()




        elif 'multiply' in query :
           speak('please enter the value of a\n')
           a = input('what is the value of a\n')
           a = int(a)    

           speak('please enter the value of b\n')
           b = input('what is the value of b\n')
           b = int(b)

           speak('a multiply b equals')
           
           speak(a * b)
           print('a * b ='+ a * b)
           engine.runAndWait()

        elif 'addition' in query :
            speak('please enter the value of a\n')
            a = input('what is the value of a\n')
            a = int(a)
 
            speak('please enter the value of b\n')
            b = input('what is the value of b\n')
            b = int(b)
 
            speak(a + b)
            print(a + b)
            engine.runAndWait()

        elif 'subtraction' in query :
           speak('please enter the value of a\n')
           a = input('what is the value of a\n')
           a = int(a)

           speak('please enter the value of b\n')
           b = input('what is the value of b\n')
           b = int(b)

           speak(a * b)
           print(a * b)
           engine.runAndWait()


        elif 'divide' in query :
           speak('please enter the value of a\n')
           a = input('what is the value of a\n')
           a = int(a)
           engine.runAndWait()

           speak('please enter the value of b\n')
           b = input('what is the value of b\n')
           b = int(b)

           print(a / b)
           speak(a / b)
           engine.runAndWait()

        elif 'square' in query :
           speak('please enter the value \n')
           a = input('what is the value \n')
           a = int(a)

           speak(a*a)
           print(a*a)
           engine.runAndWait()

 


        elif 'shut down' in query:
            speak('Sir if you need any help please ask me, i will be ready to help you')
            exit()
