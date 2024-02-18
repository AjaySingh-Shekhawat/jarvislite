import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import win32com.client


engin = pyttsx3.init('sapi5')
voices = engin.getProperty('voices')
# print(voices[1].id)
engin.setProperty('voice', voices[0].id)  # use 0 for male and 1 for female


def speak(audio):
    engin.say(audio)
    engin.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon")

    else:
        speak("Good Evening ")

    speak("I am Jarvis. How may I help you Sir.")
    # speak("boss ko sabhave jana koni k  dakho de k bara kad madarchod n ")

    # speak("I am Jarvis Sir. Please tell me how may I help you")
    # speak(" Boss ko sabhav jana koni k , uksha ro h boss n , dakho de k bara kad madarchood n ")


def takeCommand():
    # it takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"user said: {query}\n")

    except Exception as e:
        # print(e)
        print("say this again please...")
        return "None"
    return query


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gamil.com, 587')
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()


if __name__ == "__main__":
    # speak("Ajay Singh is a good boy")
    wishMe()

    while True:
        # if 1:     ####### to run only once
        query = takeCommand().lower()

        # logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('searching wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=5)
            speak("According to wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open stack overflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'play music' in query:
            music_dir = 'c:\\Users\\Genius\\Desktop\\songs\\tune'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[0]))

        elif 'tell me the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'open py' in query:
            codePath = "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2023.1.3\\bin\\pycharm64.exe"
            os.startfile(codePath)

        elif 'open valorant' in query:
            codePath = "C:\\Riot Games\\Riot Client\\RiotClientServices.exe"
            os.startfile(codePath)

        elif 'open vs' in query:
            codePath = "C:\\Users\\Genius\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'send email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "ajsinghshekhawat10@gmail.com"
                sendEmail(to, content)
                speak("Email has  been sent!")
            except Exception as e:
                print(e)
                speak("Sorry Sir, I am not able to send this mail")

        if "Jarvis quit" in query:
            exit()
