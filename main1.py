import datetime
import speech_recognition
import win32com.client
import webbrowser
import openai
import os
from config import apikey
import openpyxl

def say(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    s = text
    speaker.Speak(s)

def takeCommand():
    recognizer = speech_recognition.Recognizer()

    with speech_recognition.Microphone() as mic:
        recognizer.adjust_for_ambient_noise(source=mic, duration=0.4)
        audio = recognizer.listen(source=mic)

        try:
            text = recognizer.recognize_google(audio)
            text = text.lower()
            recognizer.pause_threshold = 0.4

            print(f"Recognized {text}")
            return text

        except:
            return("Some error ocurred")

def ai(query):
    openai.api_key = apikey
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": query}],
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    print(response["choices"][0]["message"]["content"])
    say(response["choices"][0]["message"]["content"])

'''def Attendance(text):
    workbook = openpyxl.load_workbook('workbook.xlsx')
    sheet = workbook['Sheet1']
    i=4
    while(i < 14 ):
        i = i + 1
    sheet.cell(row=i,column=1).value = 'text'  '''

if __name__ == '__main__':
    print('pycharm')
    say("Hello i am friday")
    print("Listenting....")
    text = takeCommand()
    #todo: add more sites
    sites = [["youtube","https://youtube.com"],["wikipedia","https://wikipedia.com"],["google","https://google.com"],["music","https://open.spotify.com/track/6DxVjDeLpPPimmiqSv05hd?si=42efc07c09034f25"]]

    if "play video" in text.lower():
        videopath="Videos\Captures\Obama deepfake video download - Google Search - Google Chrome 2023-08-15 17-07-32.mp4 "
        os.system(f"{videopath}")

    if "the time" in text:
        strfTime = datetime.datetime.now().strftime("%H:%M")
        say(f"Sir the time is {strfTime}")

    #todo:add more text
    if "i love you"  in text.lower():
        say(f"I love you too sir")

    if "using ai" in text.lower():
        ai(text)

    if "your name" in text.lower():
        say("my name is friday. My name is tribute to tony stark after he used this ai in avengers age of ultron")

    for site in sites:
        if f"Open {site[0]}".lower() in text.lower():
            say(f"Opening {site[0]} sir...")
            webbrowser.open(site[1])

        if f"Play {site[0]}".lower() in text.lower():
            say(f"Playing {site[0]} sir... ")
            webbrowser.open(site[1])

    if "sahil" in text.lower():
        say("Sahil is Gaddhha number1")
'''
    if "take attendance" in text.lower():
        say("ok sir taking attendance")
        Attendance(text)'''