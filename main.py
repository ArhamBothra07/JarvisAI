import speech_recognition as sr
import win32com.client
import webbrowser
import os
import datetime

def say(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(f"{text}")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 2
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            return f"User said {query}"
        except Exception as e:
            return "sorry, some error occured"

if __name__ == '__main__':
    say("Hello, this is Jarvis A,I")
    while True:
        print("Listening...")
        query = takeCommand()
        print(query)

        sites = [["youtube", "https://www.youtube.com/"], ["mail", "https://mail.google.com/mail/u/0/#inbox"], ["wikipedia", "https://www.wikipedia.org/"], ["my anime list", "https://myanimelist.net/animelist/God_D_Beast"], ["google","https://www.google.com/" ], ["9 anime", "https://9anime.pl/user/watch-list?folder=1"]]
        for site in sites:
            if f"open {site[0]}" in query.lower():
                say(f"Opening {site[0]}")
                webbrowser.open(site[1])

        apps = [["discord", "C:/Users/arham/Desktop/Discord.lnk"], ["spotify", "C:/Users/arham/Desktop/Spotify.lnk"]]
        for app in apps:
            if f"open {app[0]}" in query.lower():
                say(f"Opening {app[0]}")
                os.startfile(app[1])

        if "the time" in query.lower():
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Sir, the time right now is {strfTime}")
