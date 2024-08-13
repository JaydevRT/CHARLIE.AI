import os
import subprocess
import pyautogui
import psutil
import openai
import webbrowser
import datetime
import speech_recognition as sr
import win32com.client
from config import apikey

# Initialize voice and media player
speaker = win32com.client.Dispatch("SAPI.SpVoice")
player = win32com.client.Dispatch("WMPlayer.OCX")

chatStr = ""

def chat(query):
    global chatStr
    openai.api_key = apikey
    chatStr += f"User: {query}\nJarvis: "
    
    try:
        response = openai.Completion.create(
            model="text-davinci-003",
            prompt=chatStr,
            temperature=0.7,
            max_tokens=256,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )
        say(response["choices"][0]["text"])
        chatStr += f"{response['choices'][0]['text']}\n"
    except Exception as e:
        print(f"Error during OpenAI API call: {e}")
        return "Error occurred"

    return response["choices"][0]["text"]

def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    try:
        response = openai.Completion.create(
            model="text-davinci-003",
            prompt=prompt,
            temperature=0.7,
            max_tokens=256,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )
        text += response["choices"][0]["text"]
        if not os.path.exists("Openai"):
            os.mkdir("Openai")

        with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip() }.txt", "w") as f:
            f.write(text)
    except Exception as e:
        print(f"Error during OpenAI API call: {e}")
        return "Error occurred"

    return text

def say(text):
    speaker.Speak(text)

def close_app():
    try:
        process_list = sorted(psutil.process_iter(attrs=['pid', 'name', 'create_time']),
                              key=lambda x: x.info['create_time'], reverse=True)
        current_pid = psutil.Process().pid
        process_list = [p for p in process_list if p.info['pid'] != current_pid]

        if process_list:
            most_recent_pid = process_list[0].info['pid']
            try:
                process = psutil.Process(most_recent_pid)
                process.terminate()
                print("Most recent application has been closed.")
            except psutil.NoSuchProcess:
                print("The application is already closed.")
    except Exception as e:
        print(f"An error occurred while closing the top application: {e}")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            return query
        except Exception as e:
            print("Some Error Occurred. Sorry from Jarvis")
            return "Some Error Occurred. Sorry from Jarvis"

def process_query(query):
    sites = {
        "google": "https://www.google.com",
        "youtube": "https://www.youtube.com",
        "facebook": "https://www.facebook.com",
        "wikipedia": "https://www.wikipedia.com"
    }

    for site in sites:
        if f"open {site}".lower() in query.lower():
            say(f"Opening {site}, sir...")
            webbrowser.open(sites[site])
            say(f"{site} has been opened, sir")
        elif f"close {site}".lower() in query.lower():
            say(f"Closing {site}, sir...")
            pyautogui.hotkey('ctrl', 'w')

    if "open notepad" in query.lower():
        say("Opening notepad, sir")
        pyautogui.hotkey('win', 'r')
        pyautogui.write('notepad')
        pyautogui.press('enter')
        say("Notepad has been opened, sir")
    elif "the time" in query.lower():
        strfTime = datetime.datetime.now().strftime("%H:%M %p")
        say(f"Sir, the time is {strfTime}")
    elif "turn on wi-fi" in query.lower():
        say("Wi-Fi is now On, sir")
        subprocess.run("netsh interface set interface 'Wi-Fi' enable", shell=True)
    elif "turn off wi-fi" in query.lower():
        say("Wi-Fi is now Off, sir")
        subprocess.run("netsh interface set interface 'Wi-Fi' disable", shell=True)
    elif "close" in query.lower():
        close_app()
        pyautogui.hotkey('ctrl', 'w')
        say("Application has been closed, sir")
    elif "use ai" in query.lower():
        ai(prompt=query)
    elif "reset chat" in query.lower():
        global chatStr
        chatStr = ""
    elif "quit" in query.lower():
        say("Quitting, sir")
        exit()
    else:
        chat(query)

if __name__ == '__main__':
    print('Welcome to Jarvis A.I')
    say("Jarvis A.I")
    while True:
        query = takeCommand()
        process_query(query)
