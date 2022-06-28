import pyttsx3 #pip install pyttsx3 
#pythoncom error:pip install pywin32 
import datetime
import speech_recognition as sr #pip install SpeechRecognition
#PyAudio error:"pip install wheel" then "pip install pipwin" then "pipwin install pyaudio" If still not solved see next line
#PyAudio error:upgrade pip then goto this page-https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyaudio and download pyaudio .whl file suitable to ur pc like u had python 3.7 installed on 32 bit machine then go for file having cp37 and win32 one if on 64 bit go for amd64 one. After that open command prompt and go to the downloaded file location and use pip install {file name without brackets} then open python and run the cmd "import pyaudio" It will work 
import wikipedia #pip install wikipedia
import smtplib
import getpass
import webbrowser as wb
import os
import pyautogui #pip install pyautogui
import psutil #pip insatall psutil
import requests as rq
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from config import PROFILE_PATH




#pyttsx3 imported successfully
engine = pyttsx3.init(debug=True)  #Error:"No module pythoncom" 
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
volume = engine.getProperty('volume')
engine.setProperty('volume', volume-0.25)
rate = engine.getProperty('rate')
engine.setProperty('rate', rate-30)

def speak(audio):
	engine.say(audio)
	engine.runAndWait()

def date():
	now = datetime.datetime.now().strftime("%d %m %Y")
	speak("Today is")
	speak(now)

#for date & time install datetime
def time():
	print("in time")
	Time=datetime.datetime.now().strftime("%I %M %S")
	print(Time)
	speak("the current time is")
	print("speak the")
	speak(Time)
	print("speak Time")

def wishme():
	hour = datetime.datetime.now().hour
	if hour>=6 and hour<12:
		speak("Good Morning Sir")
	elif hour>=12 and hour<17:
		speak("Good Afternoon Sir")
	else:
		speak("Good Evening Sir")
	speak("Welcome I am Moon,ur assistant")
	date()
	speak("and")
	time()
	speak("and")
	cpu()
	speak("Moon is readily to help you")


#for user input through mic install SpeechRecognition
def takecmd():
	recog=sr.Recognizer()
	with sr.Microphone() as source: #error PyAudio
		speak("I am Listening...")
		print("Listening...")
		recog.pause_threshold=1
		audio = recog.listen(source)
	try:
		speak("I am Recognising Sir...")
		print("Recognising....")
		query = recog.recognize_google(audio,language='en-in')
		print(query)
		speak("You asked to me as ")
		speak(query)
	except Exception as e:
		print(e)
		speak("say that again please...")
		return "None"
	return(query)


def sendEmail(my_mail,pswd,to , content):
	server = smtplib.SMTP('smtp.gmail.com','587')
	server.ehlo()
	server.starttls()
	server.login(my_mail,pswd)
	server.sendmail(my_mail,to,content)
	server.close()

count=0
def screencap():
	img=pyautogui.screenshot()
	sspath = 'C://Users//HP//Desktop//Screenshots//ss'+str(count)+'.png'
	img.save(sspath)


def cpu():
	usage = str(psutil.cpu_percent())
	print("CPU is at "+usage)
	speak("CPU is at"+usage)
	battery = psutil.sensors_battery()
	print("Battery is at "+str(battery.percent)+"%")
	speak("Battery is at")
	speak(battery.percent)
	speak("%")


def weather():
	userapi = '5bba761e57b27cf86727d79466f283dc'
	location = 'Kota'
	complete_api_link = "http://api.openweathermap.org/data/2.5/weather?q="+location+"&appid="+userapi

	api_link = rq.get(complete_api_link)
	api_data = api_link.json()
	
	if api_data['cod']==404:
		print("Invalid City: {},Please Check Your City Name".format(location))
	else:
		temp = ((api_data['main']['temp'])-273.15)
		weather_desc = (api_data['weather'][0]['description'])
		humidity = (api_data['main']['humidity'])

		print("---------------------------------------------------")
		report = "Weather reports for - {}".format(location.upper())
		print(report)
		print("---------------------------------------------------")
		speak(report)

		my_temp = "Current Temperature is: {:.2f} deg C".format(temp)
		print(my_temp)
		speak(my_temp)
		
		print("Current Weather Description:",weather_desc)
		speak("Current Weather Description:"+weather_desc)

		print("Current Humidity:",humidity,'%')
		speak("Current Humidity:"+str(humidity)+'%')

def confirmation(confirm):
	speak("If you want to continue say Yes If not then say no")
	tak_confirmation=takecmd()
	return tak_confirmation


    

if __name__ == "__main__":
	TGREEN =  '\033[32m'
	TWHITE = '\033[37m'
	TYELLOW = '\033[33m'
	TBLUE = '\033[34m'
	TRED = '\033[31m'

	print(TRED,"''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''")
	print(TYELLOW,"                                                                                               AI ASSISTANT                                                                                                       ")
	print(TRED,"''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''")
	wishme()
	while True:
		speak("What u want me to do sir")
		query=takecmd()
		#query=input()
		if 'time' in query:
			time()
		elif 'date' in query:
			date()
		elif ("creator" in query ) or ("developer" in query) :
			speak("My developer is Yesh Saankhla ,he is a very Good and hard working person, I am being maintained by them and I am very happy to help them.")
		elif ("search" in query) and ("Wikipedia" in query):
			speak("What did you want to Search on Wikipedia")
			wiki_search=input()
			cnfrm=confirmation(wiki_search)
			if (cnfrm=="yes"): 
				speak("Searching...")
				query = wiki_search.replace("search" or "for","")
				result = wikipedia.summary(query,sentences=2)
				print(result)
				speak(result)
			else:
				continue
		elif ("send" in query) and (("email" in query) or ("mail" in query)):
			cnfrm=confirmation(query)
			if ((cnfrm=="yes") or (cnfrm=="Yes")):
				try:
					speak("Enter your mail id Sir")
					my_mail=input("Enter your mail id:\n")
					speak("Enter your password NOTE YOUR PASSWORD WILL NOT DISPLAY ON SCREEN")
					pswd=getpass.getpass("Enter your password:\n")
					speak("What should I mail?")
					content = input("What should I mail?\n")
					speak("To whom you want to send sir?")
					to=input("Enter Receipent Mail Address:\n")
					print("I am sending an email to",to,"and the content is",content)
					speak("I am sending an email to")
					speak(to)
					speak("and the content is")
					speak(content)
					sendEmail(my_mail,pswd,to , content)
					speak("Email has been sent to")
					speak(to)
				except Exception as e:
					print(e)
					speak("I am sorry Email was unable to sent")
			else :
				continue
		elif ('Chrome' in query) or ("Google" in query):
			speak("what should I search\n")
			cpath='C:/Program Files (x86)/Google/Chrome/Application/chrome.exe'
			search=takecmd()
			wb.get(cpath).open_new_tab(search)
		elif 'logout' in query:
			os.system("shutdown -l")
		elif 'shutdown' in query:
			os.system("shutdown /s /t 1")
		elif 'restart' in query:
			os.system("shutdown /r /t 1")
		elif ('remember that' in query) or ('make a note' in query):
			speak("What should I remember")
			data = input("Enter data to remember?")
			speak("you said me to remember "+data)
			remember = open('data.txt','a')
			remember.write(data)
			remember.close()
		elif ('do you know anything' in query) or ('have a note' in query):
			remember=open('data.txt','r')
			speak("you said me to remember that" +remember.read())
		elif 'screenshot' in query:
			try:
				count +=1
				screencap()
				speak("Screnshot sussesfully captured")
			except Exception as e:
				print(e)
				speak("Screenshot was not captured")
		elif 'cpu' in query:
			cpu()
		elif 'weather' in query:
			weather()
		elif 'turn off' in query :
			speak("okh by Yesh have a good day I will be waiting for you to meet you again ")
			quit()
		elif ("open" in query) and ("app" in query):
			speak("Which app you want to open sir")
			app_name=input("Enter the app name: \n")
			os.system(app_name)
		elif ("WhatsApp" in query):
			import whatsapp
		elif ('how are you' or 'whats up' ) in  query :
			speak("Thanks for asking I am great, working fine & happy with you also readily available for you. Hop eyou are fine also?")
			cmd = takecmd()
			if ('no' or 'not') in cmd:
				speak("Why so sir, I am here for you things would be fine lets proceed")
		else :
			speak("Sorry")

