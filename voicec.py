import subprocess
import wolframalpha
import pyttsx3
import tkinter
import calendar
import json
import random
import operator
import pyautogui
import win32com.client as win32
import speech_recognition as sr
import datetime
import wikipedia
from pytube import YouTube
import pywhatkit as kit
import webbrowser
import os
import pickle
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

import io
from PIL import Image
from stability_sdk import client
import stability_sdk.interfaces.gooseai.generation.generation_pb2 as generation



engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']


chrome_path = r'C:\Users\geeth\AppData\Local\Google\Chrome\Application\chrome.exe'
contacts = {'dharma': '+919080486995', 'yuva': '+919361317280','priya':'+916380135129'}
email_contact ={'priya':'shakthipriyaseenu@gmail.com'}


def speak(audio):
	engine.say(audio)
	engine.runAndWait()



def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<12:
		speak("Good Morning Sir !")

	elif hour>= 12 and hour<18:
		speak("Good Afternoon Sir !") 

	else:
		speak("Good Evening Sir !") 

	vname =("Capstone")
	speak("I am your Assistant")
	speak(vname)
	

def username():
	speak("What should i call you sir")
	uname = takeCommand()
	speak("Welcome Mister")
	speak(uname)
	
	speak("How can i Help you, Sir")

def takeCommand():
	
	r = sr.Recognizer()
	
	with sr.Microphone() as source:
		print("Calibrating for background noise...")
		r.adjust_for_ambient_noise(source, duration=1)
		print("Listening...")
		speak("Listening...")

		r.pause_threshold = 1
		audio = r.listen(source)

	try:
		print("Recognizing...") 
		speak("Recognizing...") 

		query = r.recognize_google(audio, language ='en-in')
		print(f"User said: {query}\n")
		speak(f"User said: {query}\n")


	except Exception as e:
		print(e) 
		print("Unable to Recognize your voice.") 
		speak("Unable to Recognize your voice.") 

		return "None"
	
	return query


def rec_audio():
    recog = sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening...")
        speak("Listening...")
        recog.pause_threshold = 1
        audio = recog.listen(source)

    data = ""

    try:
        data = recog.recognize_google(audio)
        print("You said: " + data)
        speak("You said: " + data)

    except sr.UnknownValueError:
        print("Assistant could not understand the audio")
        speak("Assistant could not understand the audio")

    except sr.RequestError as ex:
        print("Request Error from Google Speech Recognition" + ex)
        speak("Request Error from Google Speech Recognition" + ex)

    return data

def call(text):
    action_call = "assistant"

    text = text.lower()

    if action_call in text:
        return True

    return False
def say_hello(text):
    greet = ["hi", "hey", "hola", "greetings", "wassup", "hello"]

    response = ["howdy", "whats good", "hello", "hey there"]

    for word in text.split():
        if word.lower() in greet:
            return random.choice(response) + "."

    return ""

def send_whatsapp_message(name, message):
    if name in contacts:
        number = contacts[name]
        kit.sendwhatmsg_instantly(number, message)
        speak(f"Sending message to {name}")
    else:
        speak(f"Sorry, I don't have {name} in contacts")
def today_date():
    now = datetime.datetime.now()
    date_now = datetime.datetime.today()
    week_now = calendar.day_name[date_now.weekday()]
    month_now = now.month
    day_now = now.day

    months = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ]

    ordinals = [
        "1st",
        "2nd",
        "3rd",
        "4th",
        "5th",
        "6th",
        "7th",
        "8th",
        "9th",
        "10th",
        "11th",
        "12th",
        "13th",
        "14th",
        "15th",
        "16th",
        "17th",
        "18th",
        "19th",
        "20th",
        "21st",
        "22nd",
        "23rd",
        "24th",
        "25th",
        "26th",
        "27th",
        "28th",
        "29th",
        "30th",
        "31st",
    ]

    return "Today is " + week_now + ", " + months[month_now - 1] + " the " + ordinals[day_now - 1] + "."

def google_calendar():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                '../../Voice Assistant/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    return service


def sendEmail(to, content):
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.ehlo()
	server.starttls()
	server.login('shakthipriyaseenu@gmail.com', 'mtdgrhlfkpqsywwl')
	server.sendmail('shakthipriyaseenu@gmail.com', to, content)
	server.close()


def search_location(location):
    pyautogui.click(x=100, y=100) 
    time.sleep(1)
    pyautogui.write(location)
    time.sleep(1)
    pyautogui.press('enter') 
    time.sleep(5)

def navigate(start_location, destination):
    try:
        pyautogui.click(x=100, y=100) 
        time.sleep(2)
        pyautogui.click(x=200, y=200)
        pyautogui.write(start_location + '\n', interval=0.1)
        time.sleep(2) 
        pyautogui.click(x=300, y=300)
        time.sleep(2)
        pyautogui.click(x=400, y=400) 
        pyautogui.write(destination + '\n', interval=0.1)
        time.sleep(2) 
        pyautogui.click(x=500, y=500) 
        time.sleep(2)
        pyautogui.click(x=600, y=600) 
        time.sleep(5) 
        speak("Navigation started. Please follow the route.")
    except Exception as e:
        speak("An error occurred while starting navigation. Please try again later.")
        print(f"Error: {e}")

def start_google_maps():
    url = "https://www.google.com/maps"
    webbrowser.open(url)
def open_word():
    speak("Opening Microsoft Word.")
    word_app = win32.Dispatch("Word.Application")
    word_app.Visible = True 
    doc = word_app.Documents.Add()

    speak("What would you like to write?")
    text = takeCommand()
    
    if text:
        doc.Content.Text = text
        speak("Content added to the document.")
    
    while True:
        speak("Would you like to close Microsoft Word?")
        response = takeCommand()
        
        if response and "yes" in response:
            speak("Would you like to save the document?")
            save_response = takeCommand()
            
            if save_response and "yes" in save_response:
                speak("Please tell me the file name.")
                filename = takeCommand()
                
                if filename:
                    # Saving
                    doc.SaveAs(os.path.expanduser(f"~/Documents/{filename}.docx"))
                    speak(f"Document saved as {filename}.docx")
            doc.Close(SaveChanges=False)  # Close without saving if user says 'no'
            word_app.Quit()
            speak("Microsoft Word is now closed.")
            break
        elif response and "no" in response:
            speak("Microsoft Word remains open. You can continue editing or say 'close Word' to exit.")
        else:
            speak("I didn't get that. Please say 'yes' to close Word or 'no' to keep it open.")
def open_excel():
    speak("Opening Microsoft Excel.")
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True 
    workbook = excel_app.Workbooks.Add()  
    sheet = workbook.Sheets(1) 
    
    speak("What data would you like to enter in Excel?")
    #start from 1st row and col
    row = 1 
    col = 1  
    while True:
        speak(f"Please say the data for row {row}, column {col}, or say 'stop' to finish.")
        data = takeCommand()
        
        if data and data.lower() == "stop":
            speak("Data entry stopped.")
            break
        elif data:
            sheet.Cells(row, col).Value = data  # Insert data into the current cell
            speak(f"Entered '{data}' in row {row}, column {col}")
            col += 1  # Move to the next column for more data
    
    while True:
        speak("Would you like to close Microsoft Excel?")
        response = takeCommand()
        
        if response and "yes" in response:
            speak("Would you like to save the workbook?")
            save_response = takeCommand()
            
            if save_response and "yes" in save_response:
                speak("Please tell me the file name.")
                filename = takeCommand()
                
                if filename:
                    workbook.SaveAs(os.path.expanduser(f"~/Documents/{filename}.xlsx"))
                    speak(f"Workbook saved as {filename}.xlsx")
            workbook.Close(SaveChanges=False) 
            excel_app.Quit()
            speak("Microsoft Excel is now closed.")
            break
        elif response and "no" in response:
            speak("Microsoft Excel remains open. You can continue editing or say 'close Excel' to exit.")
        else:
            speak("I didn't get that. Please say 'yes' to close Excel or 'no' to keep it open.")
def open_notepad():
    speak("Opening Notepad.")
    subprocess.Popen("notepad.exe") 
    time.sleep(1) 
    
    speak("What would you like to write in Notepad?")
    while True:
        text = takeCommand()
        
        if text and "stop writing" in text:
            speak("Stopping writing in Notepad.")
            break
        elif text:
            pyautogui.write(text)
            pyautogui.press("enter")
            speak(f"Wrote '{text}' in Notepad")

    while True:
        speak("Would you like to save the Notepad file?")
        save_response = takeCommand()
        
        if save_response and "yes" in save_response:
            speak("Please tell me the file name.")
            filename = takeCommand()
            
            if filename:
                pyautogui.hotkey("ctrl", "s")
                time.sleep(1)
                pyautogui.write(f"{filename}.txt")
                pyautogui.press("enter")
                speak(f"File saved as {filename}.txt")
            break
        elif save_response and "no" in save_response:
            pyautogui.hotkey("alt", "f4")
            pyautogui.press("n")  
            speak("Notepad closed without saving.")
            break
        else:
            speak("I didn't get that. Please say 'yes' to save or 'no' to close without saving.")

			

DREAMSTUDIO_API_KEY = "sk-67uxghO02zUTjZFVnaN9e6ksjBnIZsBOQP65Vw8U2Om5C9jr"

def generate_and_save_image():
    """
    Asks for the type of image, generates the image based on the input, and saves it locally.
    """
    speak("What type of image would you like?")
    print("What type of image would you like?")
    image_type = takeCommand().lower() 

    if image_type:
        try:
            stability_api = client.StabilityInference(
                key=DREAMSTUDIO_API_KEY,
                verbose=True,
            )
            answers = stability_api.generate(
                prompt=image_type,
                seed=95456, 
            )
            for resp in answers:
                for artifact in resp.artifacts:
                    if artifact.finish_reason == generation.FILTER:
                        speak("The request triggered the API's safety filters. Please try a different prompt.")
                        print("WARNING: The request triggered the API's safety filters. Please try a different prompt.")
                        return None
                    elif artifact.type == generation.ARTIFACT_IMAGE:
                        img = Image.open(io.BytesIO(artifact.binary))

                        image_path = "generated_image.png"
                        img.save(image_path)
                        img.show()
                        
                        
                        speak("Would you like to save the image or generate another?")
                        print("Would you like to save the image or generate another?")
                        save_response = takeCommand().lower()
                        
                        if 'save' in save_response:

                            speak("Image has been saved.")
                            print("Image has been saved.")
                        elif 'generate' in save_response:
                            speak("What type of image would you like?")
                            print("What type of image would you like?")
                            image_type = takeCommand().lower() 
                            generate_and_save_image()  
                        else:
                            speak("Okay, no further action taken.")
                            print("Okay, no further action taken.")
                        return img
            return None

        except Exception as e:
            print(f"An error occurred: {e}")
            speak("An error occurred while generating the image.")
            return None
    else:
        print("No type of image specified.")
        speak("No type of image specified.")






if __name__ == '__main__':
	clear = lambda: os.system('cls')
	
	clear()
	wishMe()
	username()
	
	while True:
		query = takeCommand().lower()

		if 'wikipedia' in query:
			speak('Searching Wikipedia...')
			query = query.replace("wikipedia", "").strip()
			try:
				results = wikipedia.summary(query, sentences=3)
				speak("According to Wikipedia")
				print(results)
				speak(results)
			except wikipedia.exceptions.DisambiguationError as e:
				speak(f"Multiple results found. Can you be more specific?")
			except wikipedia.exceptions.PageError as e:
				speak(f"Sorry, I couldn't find any information about '{query}'.")

		elif call(query):
			speak(say_hello(query))	
			
		elif 'open google' in query:
			speak("What should I search?")
			search_query = takeCommand().lower()
			chrome_path = r'C:\Users\geeth\AppData\Local\Google\Chrome\Application\chrome.exe'
			webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
			webbrowser.get('chrome').open(f'https://www.google.com/search?q={search_query}')
			results = wikipedia.summary(search_query, sentences=2)
			speak(results)
		elif 'close google' in query:
			os.system("taskkill /f /im chrome.exe")
			
		elif "generate image" in query.lower() or "create image" in query.lower():
			generate_and_save_image()

		elif 'open youtube' in query:
			speak("What would you like to watch?")
			query = takeCommand().lower()
			speak(f"Playing {query} on YouTube.")
			kit.playonyt(query)
			
		elif 'close youtube' in query:
			os.system("taskkill /f /im chrome.exe")

		elif 'play music' in query:
			music_dir = 'E:\Musics'
			songs = os.listdir(music_dir) 
			os.startfile(os.path.join(music_dir, random.choice(songs)))	#######

		elif "open word" in query.lower():
			speak("Opening Microsoft Word")
			open_word()	
		elif "open excel" in query.lower():
			open_excel()
		elif "vs code" in query.lower():
			speak ("Opening Visual Studio Code")
			os.startfile(r"C:\Users\geeth\AppData\Local\Programs\Microsoft VS Code\Code.exe")
		elif "stack overflow" in query.lower():
			speak ("Opening StackOverFlow")
			webbrowser.open("https://stackoverflow.com/")
		elif "close msword" in query.lower():
			speak("Closing Microsoft Word")
			os.system("taskkill /f /im WINWORD.EXE")
		elif "close excel" in query.lower():
			speak("Closing Microsoft Excel")
			os.system("taskkill /f /im EXCEL.EXE")
		elif "close visual studio code" in query.lower():
			speak("Closing Visual Studio Code")
			os.system("taskkill /f /im Code.exe")	

		elif 'close music' in query:
			os.system("taskkill /f /im vlc.exe")	

		elif 'the time' in query:
			strTime = datetime.datetime.now().strftime("%H:%M:%S") 
			speak(f"Sir, the time is {strTime}")
			print(f"Sir, the time is {strTime}")

		elif "date" in query or "day" in query or "month" in query:
			get_today = today_date()
			print(get_today)
			speak(get_today)

		elif "what is the weather in" in query:
			key = "a2277b80672742e3ed02214ebda49974"
			weather_url = "http://api.openweathermap.org/data/2.5/weather?"
			ind = query.split().index("in")
			location = query.split()[ind + 1:]
			location = "".join(location)
			url = weather_url + "appid=" + key + "&q=" + location
			js = requests.get(url).json()
			if js["cod"] != "404":
				weather = js["main"]
				temperature = weather["temp"]
				temperature = temperature - 273.15
				humidity = weather["humidity"]
				desc = js["weather"][0]["description"]
				weatherResponse = " The temperature in Celcius is " + str(temperature) + " The humidity is " + str(
					        humidity) + " and The weather description is " + str(desc)
				speak(weatherResponse)
				print("Weather in ",location,weatherResponse)
			else:
				speak("I'm sorry,city not found")


		elif 'email to teammate' in query:
			try:
				speak("What should I say?")
				content = takeCommand()
				to = "shakthipriyaseenu@gmail.com"
				sendEmail(to, content)
				speak("Email has been sent !")
			except Exception as e:
				print(e)
				speak("I am not able to send this email")

		elif 'send a mail' in query or "send mail" in query:
			try:
				
				speak("Whom should I send it to?")
				recipient = takeCommand().lower()
				if recipient in email_contact:
					speak("What should I say?")
					content = takeCommand()
					sendEmail(email_contact[recipient], content)
					speak("Email has been sent!")
				else:
					speak("Sorry, I couldn't find the contact.")
			except Exception as e:
				print(e)
				speak("I am not able to send this email")

		elif 'add contact' in query :
			speak("What's the name of the contact?")
			contact_name = takeCommand().lower()
			speak("What's the number of the contact?")
			contact_number = takeCommand()
			contacts[contact_name] = contact_number
		elif 'send message' in query or "send a message" in query:
			speak("Who do you want to send the message to?")
			recipient = takeCommand().lower()
			if recipient in contacts:
				speak("What should I send?")
				message_content = takeCommand()
				account_sid = "AC0fd0961b92202c0d4c6c62756a0dc5e1"
				auth_token = "3f86c7066c71dd0c74cdce4c52e53c83"
				client = Client(account_sid, auth_token)
				message = client.messages.create(body=message_content, from_="+12085059794", to=contacts[recipient])
				print(message.sid)
				speak("Message sent successfully")
			else:
				speak("Sorry, I couldn't find the contact.")
		elif 'list contacts' in query:
			speak("Here are your contacts:")
			for contact, number in contacts.items():
				speak(f"{contact}: {number}")
		elif "send message" in query or "whatsapp" in query:
			speak("Who would you like to send a message to?")
			name = takeCommand().lower()
			if name:
				speak(f"What is the message for {name}?")
				message = rec_audio()
				if message:
					send_whatsapp_message(name, message)
				else:
					speak("I couldn't understand the message")

		elif 'how are you' in query:
			speak("I am fine, Thank you")
			speak("How are you, Sir")

		elif 'fine' in query or "good" in query:
			speak("It's good to know that your fine")

		elif "change my name to" in query:
			query = query.replace("change my name to", "")
			vname = query

		elif "change your name" in query:
			speak("What would you like to call me, Sir ")
			vname = takeCommand()
			speak("Thanks for naming me")

		elif "what's your name" in query or "What is your name" in query:
			vname=("Capstone")
			speak("My friends call me")
			speak(vname)
			print("My friends call me", vname)

		elif "exit" in query or "terminate" in query or "quit" in query:
			speak("Thanks for giving me your time")
			exit(0)
			
		elif 'joke' in query:
			speak(pyjokes.get_joke())


		elif "calculate" in query:
			app_id = "G74TUA-RJJAJKAAQU"
			client = wolframalpha.Client(app_id)
			ind = query.lower().split().index("calculate")
			query = query.split()[ind + 1:]
			res = client.query(" ".join(query))
			answer = next(res.results).text
			speak(f"The answer is {answer}")


		elif any(keyword in query for keyword in ["what is", "who is", "what was", "who was"]):
			app_id = "G74TUA-RJJAJKAAQU"
			client = wolframalpha.Client(app_id)
			if any(keyword in query.lower() for keyword in ["is", "was"]):
				ind = next((query.lower().split().index(keyword) for keyword in ["is", "was"] if keyword in query.lower()), None)
				if ind is not None:
					query = query.split()[ind + 1:]
					res = client.query(" ".join(query))
					try:
						answer = next(res.results).text
						speak(f"The answer is {answer}")
					except StopIteration:
						speak("Sorry, I couldn't find an answer to your question.")
				else:
					speak("I'm sorry, I couldn't understand your question.")
			else:
				speak("I'm sorry, I couldn't understand your question.")




            
		elif 'play video' in query:
			speak("Please provide the URL of the YouTube video")
			video_url = takeCommand()
			if 'youtube.com' in video_url:
				yt = YouTube(video_url)
				stream = yt.streams.get_audio_only()
				speak("Playing the requested video")
				stream.download(filename='temp')
				os.system('start temp.mp4')	

		elif "who made you" in query:
			speak("I made by students from V I T for capstone project.")


		elif 'power point presentation' in query:
			speak("opening Power Point presentation")
			power = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
			os.startfile(power)

		elif "who are you" in query:
			vname = "capstone"
			speak("I am your virtual assistant and by the way my name is " + vname)


		elif 'reason for you' in query:
			speak("I was created as a Virtual Assistant for capstone project  ")

		elif 'change the background' in query:
			ctypes.windll.user32.SystemParametersInfoW(20, 
													0, 
													"Location of wallpaper",
													0)
			speak("Background changed successfully")

		elif 'news' in query:
			try:
				apiKey = 'd0091c57e1b649ed9cf7bbe0c0da10bc'
				url = f'https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey={apiKey}'
				jsonObj = urlopen(url)
				data = json.load(jsonObj)
				i = 1
				speak('Here are some top news from The Times of India')
				print('TIMES OF INDIA\n')
				for item in data['articles']:
					print(str(i) + '. ' + item['title'] + '\n')
					print(item['description'] + '\n')
					speak(str(i) + '. ' + item['title'] + '\n')
					i += 1
			except Exception as e:
				print(str(e))

		
		elif 'lock window' in query:
				speak("locking the device")
				ctypes.windll.user32.LockWorkStation()

		elif 'shutdown system' in query:
				speak("Hold On a Sec ! Your system is on its way to shut down")
				subprocess.call('shutdown / p /f')
				
		elif 'empty recycle bin' in query:
			winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
			speak("Recycle Bin Recycled")

		elif "don't listen" in query or "stop listening" in query or "do not listen" in query:
			speak("for how many seconds do you want me to sleep")
			a = rec_audio()
			b = int(a)
			time.sleep(a)
			speak(str(a) + " seconds completed. Now you can ask me anything")



		elif "where is" in query:
			query = query.replace("where is", "")
			location = query
			speak("User asked to Locate")
			speak(location)
			webbrowser.open("https://www.google.nl/maps/place/" + location + "")
		
		elif "start" in query:
			speak("Starting Google Maps.")
			start_google_maps()
			speak("Where do you want to go?")
			destination = takeCommand()
			speak(f"Searching for {destination} on Google Maps.")
			search_location(destination)
			speak("Your destination is ready. Do you want to provide your starting location?")
			query = takeCommand()
			if "yes" in query:
				speak("Please provide your starting location.")
				start_location = takeCommand()
				speak("Starting navigation.")
				navigate(start_location, destination)
			else:
				speak("Okay, starting navigation from your current location.")
				navigate("My location", destination)
		



		elif "open camera" in query or "take a photo" in query:
			ec.capture(0, "capstone Camera ", "img.jpg")

		elif "restart" in query:
			subprocess.call(["shutdown", "/r"])
			
		elif "hibernate" in query or "sleep" in query:
			speak("Hibernating")
			subprocess.call("shutdown / h")

		elif "log off" in query or "sign out" in query:
			speak("Make sure all the application are closed before sign-out")
			time.sleep(5)
			subprocess.call(["shutdown", "/l"])

		elif "write a note" in query:
			open_notepad()
		elif "close note" in query:
			os.system("taskkill /f /im notepad.exe")

		elif "open command prompt" in query:
			os.system("start cmd")
		elif "close command prompt" in query:
			os.system("taskkill /f /im cmd.exe")
		elif "volume up" in query:
			pyautogui.press("volumeup")
			pyautogui.press("volumeup")
			pyautogui.press("volumeup")
			pyautogui.press("volumeup")
			pyautogui.press("volumeup")
			pyautogui.press("volumeup")
			pyautogui.press("volumeup")
			pyautogui.press("volumeup")
			pyautogui.press("volumeup")
			pyautogui.press("volumeup")
		
		elif "volume down" in query:
			pyautogui.press("volumedown")
			pyautogui.press("volumedown")
			pyautogui.press("volumedown")
			pyautogui.press("volumedown")
			pyautogui.press("volumedown")
			pyautogui.press("volumedown")
			pyautogui.press("volumedown")
			pyautogui.press("volumedown")
			pyautogui.press("volumedown")
			pyautogui.press("volumedown")

		elif "mute" in query:
			pyautogui.press("volumemute")

		elif "refresh" in query:
			pyautogui.moveTo(1551,551, 2)
			pyautogui.click(x=1551, y=551, clicks=1, interval=0, button='right')
			pyautogui.moveTo(1620,667, 1)
			pyautogui.click(x=1620, y=667, clicks=1, interval=0, button='left')	

		elif "scroll down" in query:
			pyautogui.scroll(1000)	

		elif "drag visual studio to the right" in query:
			pyautogui.moveTo(46, 31, 2)
			pyautogui.dragRel(1857, 31, 2)

		elif 'maximize this window' in query:
			pyautogui.hotkey('alt', 'space')
			time.sleep(1)
			pyautogui.press('x')	
		
		elif 'open new window' in query:
			pyautogui.hotkey('ctrl', 'n')
		elif 'open incognito window' in query:
			pyautogui.hotkey('ctrl', 'shift', 'n')
		elif 'minimise this window' in query:
			pyautogui.hotkey('alt', 'space')
			time.sleep(1)
			pyautogui.press('n')
		elif 'open history' in query:
			pyautogui.hotkey('ctrl', 'h')
		elif 'open downloads' in query:
			pyautogui.hotkey('ctrl', 'j')
		elif 'previous tab' in query:
			pyautogui.hotkey('ctrl', 'shift', 'tab')
		elif 'next tab' in query:
			pyautogui.hotkey('ctrl', 'tab')
		elif 'close tab' in query:
			pyautogui.hotkey('ctrl', 'w')
		elif 'close window' in query:
			pyautogui.hotkey('ctrl', 'shift', 'w')

		else:
			speak("I'm not sure about that. Can you please repeat or rephrase?")				
					
