import win32com.client as wincl
import datetime
import wikipedia
import webbrowser
import speech_recognition as sr
import os
import smtplib
from win32com.client import Dispatch

mp = Dispatch("WMPlayer.OCX")

speaker_number = 0
spk = wincl.Dispatch("SAPI.SpVoice")
vcs = spk.GetVoices()
SVSFlag = 11
print(vcs.Item(speaker_number) .GetAttribute("Name")) # speaker name
spk.Voice
spk.SetVoice(vcs.Item(speaker_number))


def speak(audio):
    spk.speak(audio)


def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour<12:
        speak("Good morning Sir!. How may I help You?")
    elif hour>=12 and hour<18:
        speak("Good Afternoon Sir!. How may I help You?")
    else:
        speak("Good Evening sir!. How may I help You?")


def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print("User said : ",query)
        except Exception as e:
            if isSleeping == 0 or isPlaying == 0:
                speak("Say that again Please...")
            return "None"
        return query


wishme()
isSleeping = 0
isPlaying = 0

while True:
    query = takecommand().lower()
    if isSleeping == 0 and isPlaying == 0:
        if "how are you" in query:
            speak("I am fine Sir")
        elif "go to sleep" in query:
            isSleeping = 1
            speak("As you wish Sir")
        elif "who are you" in query:
            speak("I am just a rather very intelligent system or JARVIS mark 1. My second version is currently under dvelopment")
        elif 'wikipedia' in query:
            speak("Searching Wikipedia")
            try:
                result = wikipedia.summary(query, sentences=2)
                speak("According to Wikipedia")
                speak(result)
            except Exception:
                speak("Sorry Sir, Result not found.")
        elif 'open youtube' in query:
            webbrowser.open("youtube.com")
        elif 'play music' in query or 'play some music' in query:
            if isPlaying == 0:
                isPlaying = 1
                music_dir = 'F:\\Music'
                songs = os.listdir(music_dir)
                os.startfile(os.path.join(music_dir, songs[0]))
        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, The time is {strTime}")
        #TODO Other functions
        else:
            pass
    else:
        if "wake up" in query and isSleeping == 1:
            isSleeping = 0
            speak("Online and ready, sir.")

        if "stop" in query and isPlaying ==1:
            mp.controls.stop()
            isPlaying = 0

