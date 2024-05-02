import pyttsx3  #Sets a engine that will speak the text
import speech_recognition as sr #defining a variable that will recognize the speech using microphone
import datetime as dt
import os #Used to open the application , i.e. the path of the application is passed as a parameter to the function , and the application is opened.
import win32com.client as wincl #Used to speak the text
import cv2 #Used to open the camera
import subprocess #Used to open the application , i.e. the path of the application is passed as a parameter to the function , and the application is opened.
import random #Used to play the music randomly , i.e. the music will be played randomly.
from requests import get #Used to get the data from the website , i.e. the data will be fetched from the website.
import wikipedia #Used to search the query in wikipedia and return the result.
import webbrowser #Used to open the browser
import pywhatkit #Used to search the query in whatsapp and return the result.
engine = pyttsx3.init() #Initializing the enginee
voices = engine.getProperty('voices') #Getting details of current voice property .
engine.setProperty('voices' , voices[0].id) #Setting up the voice property of the engine
print(voices[0].id)

#Text to Speech

def speak(audio): #Defining a function that will speak the text
    engine.say(audio) #Engine will speak the text.
    engine.runAndWait() #Engine will wait for the text to be spoken.

def takecommand():  #converts speech to text
    r = sr.Recognizer() #Initializing the recognizer
   # r.pause_threshold = 1.5 # is used to set the amount of silence (in seconds) that should be considered as the end of a speech phrase
    with sr.Microphone() as source: #Using microphone as the source
           r.adjust_for_ambient_noise(source)  #r.adjust_for_ambient_noise(source) is used to reduce the noise from the audio source , i.e. the microphone.
           print("Listening Boss...") #Printing either the microphone is listening or not
           r.pause_threshold = 1 # is used to set the amount of silence (in seconds) that should be considered as the end of a speech phrase
        #audio = r.listen(source , timeout=1 , phrase_time_limit = 5)
  # Detects speech for 1 second , if no speech is detected , 
  #it will print "Listening Boss..." and will wait for 5 seconds for the next command.
           audio = r.listen(source , timeout=None , phrase_time_limit = 5)
    
    try:          
        print("Recognizing...")
        speak("Recognizing")
        query = r.recognize_google(audio , language='en-in') #Using "Google" based for voice recognition and converting audio to text , language is set to "en-bn" i.e. English-Bangla.
        print(f"user said: {query}") #Printing the query , i.e. what the user said.
        
        
    except Exception as e:
        speak("Say that again please...")
        return takecommand()
    return query
    '''The format is
       try:
           assadsad
           asdasdassd
       except Exception as e:
           print(e)
           return "None"
    '''
def wish():
    hour = int(dt.datetime.now().hour)
    if(hour>=0 and hour <= 12):
        speak("Good Morning Sire")
        print("Good Morning Sire")
    elif(hour > 12 and hour <=18):
        speak("Good Afternoon Sire")
        print("Good Afternoon Sire")
    elif(hour > 18 and hour <=20):
        speak("Good Evening Sire")
        print("Good Evening Sire")
    elif(hour > 20 and hour <=24):
        speak("Good Night Sire")
        print("Good Night Sire")
    speak("I am Jarvis 0.1 configured by DBP. Please tell me how may I help you?")
    print("I am Jarvis 0.1 configured by DBP. Please tell me how may I help you?")
     
    

    
            

    
if __name__ == "__main__" : # function is executed from JARVIS.py file only.
    #speak("Hello Sir, This is Jarvis . How may I serve you?") #Calling the speak function and passing the text as a parameter.
    wish()
    speaker = wincl.Dispatch("SAPI.SpVoice") #Initializing the speaker , i.e. the text to speech engine.chrome
       #print("Write the text you want to speak:")
       # s = input()
       #speaker.Speak(s)
    while True:
        query = takecommand().lower() #Calling the takecommand function and converting the text to lower case s o that the command can be recognized easily.

        if "open chrome" in query:
            speak("Opening Chrome")
            print("Opening Chrome")
            npath = "C:/Program Files/Google/Chrome/Application/chrome.exe" #Path of the chrome application
            os.startfile(npath) #Opening the chrome application , using the path.
            speak("Anything else sir?")
            print("Anything else sir?")  
        
        if "close chrome" in query:
            speak("Closing Chrome")
            print("Closing Chrome")
            os.system("taskkill /im chrome.exe /f") #taskkill= used to kill the task , /im = image name , /f = force ; i.e. it will kill the task forcefully.
            speak("Anything else sir?")
            print("Anything else sir?")

        if "open notepad" in query:
            speak("Opening Notepad")
            print("Opening Notepad")
            npath = "C:/Windows/system32/notepad.exe"
            os.startfile(npath)
            speak("Anything else sir?")
            print("Anything else sir?")

        
        if "close notepad" in query:
            speak("Closing Notepad")
            print("Closing Notepad")
            os.system("taskkill /im notepad.exe /f")
            speak("Anything else sir?")
            print("Anything else sir?")

        
        if "open task manager" in query or "open task" in query:
            speak("Opening Task Manager")
            print("Opening Task Manager")
            os.system("start taskmgr")
            speak("Anything else sir?")
            print("Anything else sir?")

        
        if "close task manager" in query or "close task" in query:
            speak("Closing Task Manager")
            print("Closing Task Manager")
            os.system("taskkill /im taskmgr.exe /f")
            speak("Anything else sir?")
            print("Anything else sir?")
        
        if "open cmd" in query:
            speak("Opening Command Prompt")
            print("Opening Command Prompt")
            os.system("start cmd")
            speak("Anything else sir?")
            print("Anything else sir?")

        if "close cmd" in query:
            speak("Closing Command Prompt")
            print("Closing Command Prompt")
            os.system("taskkill /im cmd.exe /f")
            speak("Anything else sir?")
            print("Anything else sir?")

        
        if "open Geforce" in query:
            speak("Opening Geforce")
            npath = "C:/Program Files/NVIDIA Corporation/NVIDIA GeForce Experience/NVIDIA GeForce Experience"
            os.startfile(npath)
            speak("Anything else sir?")
            print("Anything else sir?")
        
        if "open telegram" in query:
            #we have to import subprocess for opening the application as it is not in the system32 folder.
            npath = "C:/Program Files/WindowsApps/TelegramMessengerLLP.TelegramDesktop_4.14.2.0_x64__t4vj0pshhgkwm/Telegram.exe"
            speak("On it sir , Opening Telegram")
            print("On it sir , Opening Telegram")
            os.startfile(npath)
            speak("Anything else sir?")
            print("Anything else sir?")

        if "close telegram" in query:
            speak("Closing Telegram")
            print("Closing Telegram")
            os.system("taskkill /im Telegram.exe /f")
            speak("Anything else sir?")
            print("Anything else sir?")

        
        if "launch whatsapp" in query:
            npath = "C:/Program Files/WhatsApp.lnk"
            speak("On it sir , Opening Whatsapp")
            print("On it sir , Opening Whatsapp")
            os.startfile(npath)
            speak("Anything else sir?")
            print("Anything else sir?")
        
        if "close whatsapp" in query:
            speak("Closing Whatsapp")
            print("Closing Whatsapp")
            os.system("taskkill /im whatsapp.exe /f")
            speak("Anything else sir?")
            print("Anything else sir?")

        if "open file manager" in query:
            speak("Which file manager do you want to open sir?")
            print("Which file manager do you want to open sir?")
            query = takecommand().lower()
            if "C drive" in query or "c drive" in query or "see drive" in query or "sea drive" in query :
                npath = "C:/"
                speak("Opening C Drive")
                print("Opening C Drive")
                os.startfile(npath)
                speak("Anything else sir?")
                print("Anything else sir?")

            if "H drive" in query or "h drive" in query or "age drive" in query or "aged drive" in query:
                npath = "H:/"
                speak("Opening H Drive")
                print("Opening H Drive")
                os.startfile(npath)
                speak("Anything else sir?")
                print("Anything else sir?")

            if "I drive"in query or "i drive" in query or "eye drive" in query or "ay drive" in query:
                npath = "I:/"
                speak("Opening I Drive")
                print("Opening I Drive")
                os.startfile(npath)
                speak("Anything else sir?")
                print("Anything else sir?")

            else:
                speak("Sorry sir , I can't find the drive you are looking for")
                print("Sorry sir , I can't find the drive you are looking for")
                speak("Anything else sir?")
                print("Anything else sir?")

        if "close file manager" in query:
            speak("Closing File Manager")
            print("Closing File Manager")
            os.system("taskkill /im explorer.exe /f")
            speak("Anything else sir?")
            print("Anything else sir?")
        
        if "open black hole" in query:
            speak("Opening Black hole")
            print("Opening Black hole")
            npath="C:/Users/panth/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/BlackHole.lnk"
            os.startfile(npath)
            speak("Anything else sir?")
            print("Anything else sir?")
        
        
        #Opening Camera
            
        if "open camera" in query:
                speak("Opening Camera")
                print("Opening Camera")
                capture = cv2.VideoCapture(0) #0 means the default camera , 1 means the external camera . Here capture is the variable that will store the image.
                while True:
                    ret , image = capture.read() #ret means return , i.e. it will return the image , image is the variable that will store the capture image.
                    cv2.imshow('webcam' , image) #cv2.imshow is used to show the image , 'Camera' is the name of the window , image is the variable that will store the capture image.
                    k = cv2.waitKey(50) #cv2.waitKey is used to wait for the key to be pressed , 50 is the delay in milliseconds.
                    if k == 27:  #27 is the ASCII value of the escape key , i.e. if the escape key is pressed , the camera will close.
                      break
            
                capture.release() #capture.release() is used to release the camera.
                cv2.destroyAllWindows() #cv2.destroyAllWindows() is used to destroy all the windows.
            #both the functions are used to close the camera and outside the while loop , so that the camera will close after the escape key is pressed.
                speak("Anything else sir?")
                print("Anything else sir?")
       
    #playing music
        if "play music" in query:
            speak("Playing Music")
            print("Playing Music")
            music_dir = "H:\\MUSIC" #Path of the music folder , // is used to avoid the error.
            songs  = os.listdir(music_dir) 
            #os.listdir is used to list all the songs in the music folder.
            rd = random.choice(songs)
            #os.startfile(os.path.join(music_dir , rd)) 
            os.startfile(os.path.join(music_dir , songs[3]))
            #os.path.join is used to join the path of the music folder and the random song that is selected.
            #os.path = C://Music , os.path.join = 
            speak("Anything else sir?")
            print("Anything else sir?")
                


    #### NOW ONLINE TASKS
        if "ip address" in query or "ip" in query or "IP address" in query:
           ip = get('https://api.ipify.org').text 
      #get is used to get the data from the website , i.e. the data will be fetched from the website.
      # .text is used to convert the data to text.
           speak(f"Your IP Address is {ip}")
           print(f"Your IP Address is {ip}")
           query = query.replace("ip address", " ").replace("ip" , " ").replace("IP address" , " ")
            #Here f is used to format the string when we are using variables inside the string.
            #AND the variable is inside the curly braces.
           speak("Anything else sir?")
           print("Anything else sir?")


        
        if "wikipedia" in query or "Wikipedia" in query:
            speak("searching wikipedia...")

            #query = query.replace("wikipedia", "") #Replacing the word wikipedia with blank space , becaue wikipedia is not needed in the query.
               #This could be useful in a scenario where the user asks Jarvis to look up something on Wikipedia. 
               #For example, if the user says "Jarvis, search Wikipedia for Python", 
               #the query string would initially be "search Wikipedia for Python".
               #By removing the word "Wikipedia", the query string becomes "search for Python", 
               #which might be easier to handle in the rest of the code.
            try:                 #wikipedia.summary is used to search the query in wikipedia and sentences=5 means it will return 5 sentences.
                results = wikipedia.summary(query , sentences=2)
                speak("according to wikipedia")
                print(results)
                speak(results)
                speak("Anything else sir?")
                print("Anything else sir?")

        
            except wikipedia.DisambiguationError as e: 
                #DisambiguationError is used to handle the error ,
                #i.e. if the query is not found in wikipedia , it will show the error.
                speak("Can you please be more specific? There are multiple matches for your search.")
                print("Can you please be more specific? There are multiple matches for your search.")
            except Exception as e:
                #the second exception is used to handle the error , i.e. if the query is not found in wikipedia , 
                #it will show the error.
                speak("Sorry sir , I can't find the query you are looking for")
                print("Sorry sir , I can't find the query you are looking for")

        
        #Searching specific query in google
        elif "open google" in query or "Open google" in query:
            speak("What do you want to search sir?")
            print("What do you want to search sir?")
            search = takecommand().lower()
            speak("Opening" + search + "in Google")
            print("Opening" + search + "in Google")
            webbrowser.open(f"{search}")
            #webbrowser.open is used to open the browser , and the link is passed as a parameter.
            #?q= is used to search the query in google. 
            #"q" not "query" because it is the parameter that google uses to search the query.


            speak("Anything else sir?")
            print("Anything else sir?")

        elif "open youtube" in query or "Open youtube" in query:
            speak("Opening Youtube")
            print("Opening Youtube")
            webbrowser.open("www.youtube.com")
            #webbrowser.open is used to open the browser , and the link is passed as a parameter.
            speak("Anything else sir?")
            print("Anything else sir?")
        
        elif "open classroom" in query or "Open classroom" in query:
            speak("Opening Classroom")
            print("Opening Classroom")
            webbrowser.open("https://www.classroom.google.com")
            speak("Anything else sir?")
            print("Anything else sir?")
        
        elif "open facebook" in query or "Open facebook" in query:
            speak("Opening Facebook")
            print("Opening Facebook")
            webbrowser.open("www.facebook.com")
            speak("Anything else sir?")
            print("Anything else sir?")
        
        if "stop" in query:
            speak("Thank you sir , Have a nice day")
            print("Thank you sir , Have a nice day")
            break
        


        
'''
engine.setProperty('voices',voices[0].id)

def dpeak(audio):
    engine.say(audio)
    engine.runAndWait()

if __name__ == "__main__":
'''
