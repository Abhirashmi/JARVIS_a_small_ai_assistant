'''J.A.R.V.I.S. (Just A Rather Very Intelligent System) is a fictional artificial intelligence that first
appeared in the Marvel Cinematic Universe where he was voiced by Paul Bettany in Iron Man, Iron Man 2,
The Avengers, Iron Man 3, and Avengers: Age of Ultron.

by-ABHIRASHMI
DATE-4 th january 2021

'''
import os
import datetime
import speech_recognition as sr
import pyttsx3
from pyttsx3 import voice
import wikipedia
import webbrowser
import random
import smtplib
import requests
import json

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices)
engine.setProperty('voice', voices[0].id)


# engine.setProperty('voice',voices[1].id)
# print(voices[0].id)
# print(voices[1].id)
def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wish_me():
    speak1(
        f"hello ,{user}  i am jarvis   J.A.R.V.I.S. (Just A Rather Very Intelligent System) ")
    hour = int(datetime.datetime.now().hour)
    print(hour)
    if (hour >= 4 and hour < 11):
        speak1("Good morning")
    elif (hour >= 11 and hour < 17):
        speak1("Good afternoon")
    elif (hour >= 17 and hour <= 21):
        speak1("Good evening")
    else:
        speak1(f"why so serious human!!! its already {hour}...")
    speak1(f"How may i help you {user}")


def take_command():
    # it takes microphone input from the user and returns string output.
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print(f"Listening .........say now ")
        speak1("listening ...........say now ")
        r.pause_threshold = 1
        audio = r.listen(source)

        try:
            print("Recognizing.........PLease wait")
            speak1("Recognizing.........PLease wait")

            query = r.recognize_google(audio, language='en-in')
            print(f"{user}  said: {query} \n")
            speak1(f"{user}  said: {query} \n")

        except Exception as e:
            # print(e)
            print("Say that again please ")
            speak1("Say that again please ")
            return "None"


        return query


def speak1(str):
    from win32com.client import Dispatch
    speak1 = Dispatch("SAPI.SpVoice")
    speak1.Speak(str)

# def sendmail(to,content):
#     server=smtplib.SMTP('smtp.gmail.com',587)
#     server.ehlo()
#     server.starttls()
#     server.login('abhirashmi2000@gmail.com',password="")
#     server.sendmail()
#     server.close()


if __name__ == '__main__':
    speak1("Enter  your  name  user ")
    user = input("ENTER YOUR NAME USER ?")
    wish_me()
    if 1:
        query = take_command().lower()
        # logic to execute task based on query
        if query=="none":
            speak1(f"say something {user}")

        else:
            if "wikipedia" in query:
                speak("searching wikipedia")
                query = query.replace("wikipedia", " ")
                results = wikipedia.summary(query, sentences=3)
                speak(f"According to wikipedia .......{results}")
                print(results)


            elif "youtube" in query:
                webbrowser.open("youtube.com")

            elif "whatsapp" in query:
                webbrowser.open("whatsapp.com")

            elif "google" in query:
                webbrowser.open("google.com")

            elif "song" in query:
                webbrowser.open("https://www.youtube.com/watch?v=r-pKo1tZ52w")


            elif "news" in query:
                speak1("News for today.. Lets begin")
                url = "https://newsapi.org/v2/top-headlines?sources=the-times-of-india&apiKey" \
                      "=d093053d72bc40248998159804e0e67d "
                news = requests.get(url).text
                news_dict = json.loads(news)
                arts = news_dict['articles']
                for article in arts:
                    speak1(article['title'])
                    print(article['title'])
                    speak1("Moving on to the next news.......Listen Carefully")

                speak1("Thanks for listening...")

            elif "weather" in query:
                webbrowser.open("google.com/weather")

            elif "stack overflow" in query:
                webbrowser.open("stackoverflow.com")

            elif "play music" in query:
                music_dir="C:\\Users\\KIIT\\Desktop\\songs"
                songs=os.listdir(music_dir)
                print(songs)
                r=random.randint(0,8)
                os.startfile(os.path.join(music_dir,songs[r]))


            elif "time" in query:
                str_Time=datetime.datetime.now().strftime("%H:%M:%S")
                speak1(f"THE TIME IS {str_Time}")

            elif "vs code" in query:
                vs_code_path="ata\\Local\\Programs\\Microsoft VS Code\\Code.exe"
                os.startfile(vs_code_path)

            elif "pycharm" in query:
                pycharm_code_path="C:\Program Files\JetBrains\PyCharm Community Edition 2020.2.4\bin\pycharm64.exe"
                os.startfile(pycharm_code_path)

            elif "bye" in query:
                quit()



            # # elif "email" in query:
            #     try:
            #         speak1("what should i say ?")
            #         content=take_command()
            #         to='1828223@kiit.ac.in'
            #         sendmail(to,content)
            #         speak1("email has been sent succesfully ")

                # except Exception as e:
                #     print(e)
                #     speak1(f"sorry something went wrong {user}")

            else:
                webbrowser.open(f"{query}google search.com")

