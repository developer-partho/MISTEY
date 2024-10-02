import tkinter as tk
from tkinter import PhotoImage
from tkinter.scrolledtext import ScrolledText
import threading
import queue
import time
import os
import random
import datetime
import smtplib
import subprocess
import wikipedia
import pyjokes
import webbrowser
import requests
import wolframalpha
import pyautogui
from typing import Callable
import mouse
import shutil


from pynput.keyboard import Controller as KeyboardController
from pynput import keyboard
import speech_recognition as sr
import pyttsx3
from config import apikey
from openai import OpenAI
import re
import ctypes
import win32gui
import win32con
import screen_brightness_control as sbc
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

# ----------------------- Configuration -----------------------

# Replace with your WolframAlpha App ID
WOLFRAMALPHA_APP_ID = "YOUR_WOLFRAMALPHA_APP_ID"

# Email credentials (Use environment variables or secure methods in production)
EMAIL_ADDRESS = 'your_email@gmail.com'
EMAIL_PASSWORD = 'your_password'

# Default recipient email
DEFAULT_RECIPIENT = 'recipient_email@gmail.com'

# Paths to icons (Ensure these files exist in the same directory or provide full paths)
LISTEN_ICON_PATH = "listen_icon.png"
RECOGNIZE_ICON_PATH = "recognize_icon.png"

# Music directory path
MUSIC_DIR = "P:\\MUSIC\\Music Compositions & Mixings\\MY SONGS ALBAM"

# -------------------------------------------------------------

# Initialize Text-to-Speech Engine
engine = pyttsx3.init()
#-----------------------------------Open Ai-------------------------------------------------
client = OpenAI(
  api_key=  apikey# this is also the default, it can be omitted
)

chatStr = ""
# https://youtu.be/Z3ZAJoi4x6Q
def chat(query):
    global chatStr
    print(chatStr)
    OpenAI.api_key = apikey
    chatStr += f"Harry: {query}\n Jarvis: "
    response = OpenAI.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # todo: Wrap this inside of a  try catch block
    speak(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]

def ai(prompt):
    OpenAI.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    response = OpenAI.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # todo: Wrap this inside of a  try catch block
    # print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    # with open(f"Openai/prompt- {random.randint(1, 2343434356)}", "w") as f:
    with open(f"Openai/{''.join(prompt.split('brain')[1:]).strip()}.txt", "w") as f:
        f.write(text)


#--------------------------------------------Take Command--------------------------------------------
def speak(text):
    """
    :param text: The text string that the text-to-speech engine will convert to speech.
    :return: None
    """
    engine.say(text)
    engine.runAndWait()


# Initialize Speech Recognizer
recognizer = sr.Recognizer()


# Function to change voice (male/female)
def set_voice(voice_gender="female"):
    voices = engine.getProperty('voices')

    if voice_gender == "male":
        engine.setProperty('voice', voices[0].id)  # Set to male voice
    else:
        engine.setProperty('voice', voices[1].id)  # Set to female voice


# Optional: Adjust speech rate and volume
def set_rate_and_volume(rate=150, volume=1.0):
    engine.setProperty('rate', rate)  # Speed of speech
    engine.setProperty('volume', volume)  # Volume level (0.0 to 1.0)


# Set voice and speech rate
set_voice(voice_gender="female")  # Change to "male" for male voice
set_rate_and_volume(rate=180, volume=1.0)

def take_command(q, ui_queue):
    """
    :param q: Queue object used for inter-thread communication (not utilized in the current function implementation).
    :param ui_queue: Queue object used to update the UI thread with status and recognized user input.
    :return: This function runs indefinitely, listening for user's voice commands, and processing the recognized commands.
    """
    while True:
        try:
            with sr.Microphone() as source:
                ui_queue.put(("status", "listening"))
                print("Listening...")

                recognizer.adjust_for_ambient_noise(source)
                audio = recognizer.listen(source)

                ui_queue.put(("status", "recognizing"))
                print("Recognizing...")

            query = recognizer.recognize_google(audio, language='en-in')
            print(f"User said: {query}\n")
            ui_queue.put(("user_input", query))

        except Exception as e:
            print("Unable to Recognize your voice.")
            ui_queue.put(("system_output", "Unable to Recognize your voice."))
            continue

        process_command(query.lower(), ui_queue)

def wishme():
    """
    Evaluates the current hour and provides an appropriate greeting based on the time of day.
    Also announces the bot's name as "Misty" and sends relevant messages to a UI queue.

    :return: None
    """
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 6:
        speak("Welcome at Mid night !")
        ui_queue.put(("system_output", "Welcome at Mid night !."))
    elif 6 <= hour < 12:
        speak("Good Morning !")
        ui_queue.put(("system_output", "Good Morning !."))
    elif 12 <= hour < 3:
        speak("Good Noon !")
        ui_queue.put(("system_output", "Good Noon !."))
    elif 12 <= hour < 18:
        speak("Good Afternoon !")
        ui_queue.put(("system_output", "Good Afternoon !."))
    else:
        speak("Good Evening!")
        ui_queue.put(("system_output", "Good Evening !."))
    botname = "Misty"
    speak("I am your Assistant"+ botname)

def send_email(send_to, content):
    """
    :param send_to: The email address of the recipient.
    :param content: The content of the email to be sent.
    :return: None
    """
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, send_to, content)
        server.close()
        print("Email has been sent!")
        speak("Email has been sent!")
        ui_queue.put(("system_output", "Email has been sent!"))
    except Exception as e:
        print("I am not able to send this email.")
        speak("I am not able to send this email.")
        ui_queue.put(("system_output", "I am not able to send this email."))


def terminate_process(process_name):
    """
    :param process_name: The name of the process to terminate.
    :return: None
    """
    try:
        os.system(f'taskkill /IM "{process_name}" /F')
        print(f"{process_name} terminated.")
        speak(f"{process_name} terminated.")
        ui_queue.put(("system_output", f"{process_name} terminated."))
    except Exception as e:
        print(f"Failed to terminate {process_name}.")
        speak(f"Failed to terminate {process_name}.")
        ui_queue.put(("system_output", f"Failed to terminate {process_name}."))

#-----------------------------------------Screen Mouse____________________________________________________

# Set global variable to track if the screen is active
mouse_screen_active = False


# Function to create a transparent window with grid
def create_grid(canvas, width, height, spacing=200):
    # Draw vertical grid lines with labels for X coordinates
    for x in range(0, width, spacing):
        canvas.create_line(x, 0, x, height, fill="white", dash=(2, 4))
        canvas.create_text(x + 10, 10, text=f"X={x}", fill="white", anchor="nw", font=("Arial", 12))

    # Draw horizontal grid lines with labels for Y coordinates
    for y in range(0, height, spacing):
        canvas.create_line(0, y, width, y, fill="white", dash=(2, 4))
        canvas.create_text(10, y + 10, text=f"Y={y}", fill="white", anchor="nw", font=("Arial", 12))


# Function to update the mouse position dynamically
def update_mouse_position():
    global mouse_screen_active
    if not mouse_screen_active:  # Stop updating if the screen is not active
        return

    # Get the current mouse position
    x, y = pyautogui.position()
    # Update the mouse position label
    mouse_position_label.config(text=f"Mouse Position: X={x}, Y={y}")
    # Repeat the update every 100 ms
    root.after(100, update_mouse_position)


# Function to make the window click-through using the Windows API
def make_window_clickthrough(hwnd):
    # Get the current window styles
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    # Add WS_EX_LAYERED and WS_EX_TRANSPARENT to allow click-through
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style | win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT)
    # Set the transparency level to 128 (50% opacity)
    win32gui.SetLayeredWindowAttributes(hwnd, 0, 128, win32con.LWA_ALPHA)


# Function to remove click-through (restore normal behavior)
def remove_clickthrough(hwnd):
    # Get the current window styles
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    # Remove WS_EX_TRANSPARENT to disable click-through
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style & ~win32con.WS_EX_TRANSPARENT)


# Function to activate the screen that shows mouse coordinates
def open_grid_screen():
    global root, canvas, mouse_position_label, mouse_screen_active
    if not mouse_screen_active:
        mouse_screen_active = True

        # Create the main Tkinter window
        root = tk.Tk()

        # Set the window to be transparent and full screen
        root.attributes("-fullscreen", True)
        root.attributes("-alpha", 0.5)  # Set transparency level
        root.attributes("-topmost", True)  # Keep the window on top
        root.configure(bg='black')

        # Create a canvas to draw the grid
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        canvas = tk.Canvas(root, width=screen_width, height=screen_height, bg='black', highlightthickness=0)
        canvas.pack(fill=tk.BOTH, expand=True)


        # Call function to create the grid with X and Y coordinates
        create_grid(canvas, screen_width, screen_height, spacing=100)

        # Create a label to display the current mouse position
        global mouse_position_label
        mouse_position_label = tk.Label(root, text="", font=("Arial", 24), fg="white", bg="black")
        mouse_position_label.pack(pady=20)

        # Start updating the mouse position dynamically
        update_mouse_position()

        # Wait for Tkinter window to fully initialize
        root.update_idletasks()

        # Get the HWND (handle) of the Tkinter window after it's created
        hwnd = ctypes.windll.user32.FindWindowW(None, root.wm_title())
        if hwnd:
            # Make the window click-through using the correct HWND
            make_window_clickthrough(hwnd)
        else:
            print("Failed to retrieve HWND for Tkinter window.")

        # Bind the Escape key to exit full-screen mode
        root.bind("<Escape>", lambda e: close_grid_screen())

        # Run the Tkinter event loop
        root.mainloop()


# Function to close the screen that shows mouse coordinates
def close_grid_screen():
    global root, mouse_screen_active
    mouse_screen_active = False
    if root:
        hwnd = ctypes.windll.user32.FindWindowW(None, root.wm_title())
        if hwnd:
            # Restore the window's normal behavior by removing click-through
            remove_clickthrough(hwnd)
        root.destroy()
#------------------------------------------Flexibility Command definition------------------------------------------------
def get_level_from_query(query):
    """Extracts a level (0-100) from the query string."""
    level = ''.join(filter(str.isdigit, query))
    return int(level) if level else 50  # Default to 50 if no number is found


def set_volume(volume_level):
    try:
        if not (0 <= volume_level <= 100):
            raise ValueError("Volume level must be between 0 and 100.")

        volume = volume_level / 100.0
        ctypes.windll.winmm.waveOutSetVolume(0, int(volume * 0xFFFF))
        ui_queue.put(f"Volume set to {volume_level}%.")

    except Exception as e:
        print(f"Error setting system volume: {e}")
        ui_queue.put(f"Error setting system volume: {e}")

    except Exception as e:
        print(f"Error setting volume: {e}")
        ui_queue.put(f"Error setting volume: {e}")

############################Command Area!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
def process_command(query, ui_queue):
    """
    :param query: The command or query given by the user to the assistant.
    :param ui_queue: A queue to communicate with the UI thread, typically used to update UI elements or send system output messages.
    :return: None
    """
    if 'wikipedia' in query:
        speak('Searching Wikipedia...')
        ui_queue.put(("system_output", "Searching Wikipedia..."))
        query = query.replace("wikipedia", "")
        try:
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            print(results)
            ui_queue.put(("system_output", results))
            speak(results)
        except Exception as e:
            speak("Sorry, I couldn't find any results on Wikipedia.")
            ui_queue.put(("system_output", "Sorry, I couldn't find any results on Wikipedia."))

    elif 'open youtube' in query:
        speak("Here you go to Youtube")
        ui_queue.put(("system_output", "Opening YouTube..."))
        webbrowser.open("https://youtube.com")
        #-----------------------------------------------------------------------open ai----------------------
    elif "Using brain".lower() in query.lower():
        ai(prompt=query)
    # -----------------------------------------------------------------------open ai----------------------
    elif 'open google' in query:
        speak("Here you go to Google")
        ui_queue.put(("system_output", "Opening Google..."))
        webbrowser.open("https://google.com")

    elif 'play music' in query or "play song" in query:
        speak("Playing music")
        ui_queue.put(("system_output", "Playing music..."))
        try:
            songs = os.listdir(MUSIC_DIR)
            if songs:
                song = random.choice(songs)
                os.startfile(os.path.join(MUSIC_DIR, song))
            else:
                speak("No songs found in the directory.")
                ui_queue.put(("system_output", "No songs found in the directory."))
        except Exception as e:
            speak("Unable to play music.")
            ui_queue.put(("system_output", "Unable to play music."))

    elif 'time now' in query or 'the time' in query:
        current_time = datetime.datetime.now().strftime('%I:%M %p')
        speak(f'Current time is {current_time}')
        ui_queue.put(("system_output", f'Current time is {current_time}'))

    elif 'date today' in query:
        today = datetime.datetime.now().strftime('%d %B %Y')
        speak(f"Today's date is {today}")
        ui_queue.put(("system_output", f"Today's date is {today}"))

    elif 'email to' in query:
        try:
            speak("What should I say?")
            ui_queue.put(("system_output", "What should I say?"))
            content = take_command()
            to = DEFAULT_RECIPIENT
            send_email(to, content)
        except Exception as e:
            speak("I am not able to send this email.")
            ui_queue.put(("system_output", "I am not able to send this email."))

    elif 'send a mail' in query:
        try:
            speak("What should I say?")
            ui_queue.put(("system_output", "What should I say?"))
            content = take_command()
            speak("Whom should I send")
            ui_queue.put(("system_output", "Whom should I send"))
            to = take_command()
            send_email(to, content)
        except Exception as e:
            speak("I am not able to send this email.")
            ui_queue.put(("system_output", "I am not able to send this email."))

    elif 'how are you' in query:
        speak("I am fine, Thank you. How are you, user?")
        ui_queue.put(("system_output", "I am fine, Thank you. How are you, Sir?"))

    elif 'fine' in query or "good" in query:
        speak("It's good to know that you're fine.")
        ui_queue.put(("system_output", "It's good to know that you're fine."))

    elif "change my name to" in query:
        global bot_name
        bot_name = query.replace("change my name to", "").strip()
        speak(f"Okay, I will call myself {bot_name} from now on.")
        ui_queue.put(("system_output", f"Changed name to {bot_name}."))

    elif "change your name" in query:
        speak("What would you like to call me, Sir?")
        ui_queue.put(("system_output", "What would you like to call me, Sir?"))
        bot_name = take_command()
        speak(f"Thanks for naming me {bot_name}.")
        ui_queue.put(("system_output", f"Thanks for naming me {bot_name}."))

    elif "what's your name" in query or "what is your name" in query:
        speak(f"My friends call me {bot_name}.")
        ui_queue.put(("system_output", f"My friends call me {bot_name}."))

    elif 'exit' in query or 'quit' in query or 'bye' in query:
        speak("Thanks for giving me your time.")
        ui_queue.put(("system_output", "Exiting..."))
        os._exit(0)

    elif "who made you" in query or "who created you" in query:
        speak("I have been created by Partho Protim Roy, a student of B.Tech CSE at Sharda University.")
        ui_queue.put(("system_output",
                      "I have been created by Partho Protim Roy, a student of B.Tech CSE at Sharda University."))

    elif 'joke' in query:
        joke = pyjokes.get_joke()
        speak(joke)
        ui_queue.put(("system_output", joke))

    elif "calculate" in query:
        try:
            speak("Calculating...")
            ui_queue.put(("system_output", "Calculating..."))
            app_id = WOLFRAMALPHA_APP_ID
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            calc_query = ' '.join(query.split()[indx + 1:])
            res = client.query(calc_query)
            answer = next(res.results).text
            print(f"The answer is {answer}")
            speak(f"The answer is {answer}")
            ui_queue.put(("system_output", f"The answer is {answer}"))
        except Exception as e:
            speak("I couldn't calculate that.")
            ui_queue.put(("system_output", "I couldn't calculate that."))

    elif 'search' in query or 'play' in query or "stars" in query:
        query = query.replace("search", "").replace("play", "").replace("stars", "").strip()
        speak(f"Searching for {query}")
        ui_queue.put(("system_output", f"Searching for {query}"))
        webbrowser.open(f"https://www.google.com/search?q={query}")

    elif "who am i" in query:
        speak("If you can talk and think, then you are a human.")
        ui_queue.put(("system_output", "If you can talk and think, then you are a human."))

    elif "why you have been created" in query:
        speak("Thanks to Partho, I am here to help people and students use technology more efficiently.")
        ui_queue.put(("system_output",
                      "Thanks to Partho, I am here to help people and students use technology more efficiently."))

    elif 'is love' in query:
        speak("Love is the 7th sense that destroys all other senses.")
        ui_queue.put(("system_output", "Love is the 7th sense that destroys all other senses."))

    elif "who are you" in query:
        speak("I am your virtual assistant created by Partho Protim Roy.")
        ui_queue.put(("system_output", "I am your virtual assistant created by Partho Protim Roy."))

    elif 'reason for you' in query:
        speak("I have been created as a minor project by Mister Partho Protim Roy.")
        ui_queue.put(("system_output", "I have been created as a minor project by Mister Partho Protim Roy."))

    elif 'change background' in query:
        speak("Changing the background.")
        ui_queue.put(("system_output", "Changing the background."))
        # Provide the path to your wallpaper image
        wallpaper_path = "path_to_your_wallpaper.jpg"
        ctypes.windll.user32.SystemParametersInfoW(20, 0, wallpaper_path, 0)
        speak("Background changed successfully.")
        ui_queue.put(("system_output", "Background changed successfully."))

    elif 'open mouse' in query or 'open virtual mouse' in query:
        try:
            mouse_app_path = r"C:\Path\To\Your\mouse_application.py"
            os.startfile(mouse_app_path)
            speak("Opening virtual mouse.")
            ui_queue.put(("system_output", "Opening virtual mouse."))
        except Exception as e:
            speak("Unable to open virtual mouse.")
            ui_queue.put(("system_output", "Unable to open virtual mouse."))

    elif 'news' in query:
        try:
            api_key = "YOUR_NEWSAPI_KEY"
            url = f'https://newsapi.org/v2/top-headlines?country=us&apiKey={api_key}'
            response = requests.get(url)
            data = response.json()
            articles = data.get('articles', [])
            if articles:
                speak('Here are some top news from the News API')
                ui_queue.put(("system_output", "Here are some top news from the News API"))
                for i, article in enumerate(articles[:5], 1):
                    title = article.get('title', 'No Title')
                    description = article.get('description', 'No Description')
                    news = f"{i}. {title}\n{description}"
                    print(news)
                    speak(news)
                    ui_queue.put(("system_output", news))
            else:
                speak("No news available at the moment.")
                ui_queue.put(("system_output", "No news available at the moment."))
        except Exception as e:
            speak("Sorry, I couldn't fetch the news.")
            ui_queue.put(("system_output", "Sorry, I couldn't fetch the news."))

    elif 'lock window' in query:
        speak("Locking the device.")
        ui_queue.put(("system_output", "Locking the device."))
        ctypes.windll.user32.LockWorkStation()

    elif 'shutdown system' in query:
        speak("Hold On a Sec! Your system is shutting down.")
        ui_queue.put(("system_output", "Shutting down the system."))
        subprocess.call('shutdown /s /t 1')

    elif 'restart system' in query:
        speak("Hold On a Sec! Your system is restarting.")
        ui_queue.put(("system_output", "Restarting the system."))
        subprocess.call('shutdown /r /t 1')

    elif "hibernate" in query or "go to sleep" in query:
        speak("Hibernating.")
        ui_queue.put(("system_output", "Hibernating."))
        subprocess.call("shutdown /h")

    elif "log off" in query or "sign out" in query:
        speak("Signing out.")
        ui_queue.put(("system_output", "Signing out."))
        subprocess.call(["shutdown", "/l"])

    elif 'empty recycle bin' in query:
        try:
            import winshell
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin emptied.")
            ui_queue.put(("system_output", "Recycle Bin emptied."))
        except Exception as e:
            speak("Unable to empty Recycle Bin.")
            ui_queue.put(("system_output", "Unable to empty Recycle Bin."))

    elif "don't listen" in query or "stop listening" in query or "do not listen" in query:
        try:
            speak("For how many seconds do you want me to stop listening?")
            ui_queue.put(("system_output", "For how many seconds do you want me to stop listening?"))
            a = int(take_command())
            speak(f"Okay, I will not listen for {a} seconds.")
            ui_queue.put(("system_output", f"Will not listen for {a} seconds."))
            time.sleep(a)
        except:
            speak("I didn't catch that.")
            ui_queue.put(("system_output", "I didn't catch that."))

    elif "where is" in query:
        query = query.replace("where is", "").strip()
        location = query
        speak(f"User asked to locate {location}")
        ui_queue.put(("system_output", f"Locating {location}"))
        webbrowser.open(f"https://www.google.nl/maps/place/{location}")
        pyautogui.press('down')
        pyautogui.press('enter')

    elif "take a photo" in query:
        try:
            import cv2
            speak("Taking a photo.")
            ui_queue.put(("system_output", "Taking a photo."))
            cam = cv2.VideoCapture(0)
            ret, frame = cam.read()
            if ret:
                cv2.imwrite("img.jpg", frame)
                speak("Photo taken and saved as img.jpg.")
                ui_queue.put(("system_output", "Photo taken and saved as img.jpg."))
            cam.release()
            cv2.destroyAllWindows()
        except Exception as e:
            speak("Unable to access the camera.")
            ui_queue.put(("system_output", "Unable to access the camera."))

    elif "write a note" in query:
        try:
            speak("What should I write, sir?")
            ui_queue.put(("system_output", "What should I write, sir?"))
            note = take_command()
            with open('mistey.txt', 'w') as file:
                speak("Sir, should I include date and time?")
                ui_queue.put(("system_output", "Should I include date and time?"))
                snfm = take_command()
                if 'yes' in snfm or 'sure' in snfm:
                    str_time = datetime.datetime.now().strftime("%I:%M %p")
                    file.write(f"{str_time} :- {note}")
                else:
                    file.write(note)
            speak("I've made a note for you.")
            ui_queue.put(("system_output", "I've made a note for you."))
        except Exception as e:
            speak("I am not able to write the note.")
            ui_queue.put(("system_output", "I am not able to write the note."))

    elif "show the note" in query:
        try:
            with open("mistey.txt", "r") as file:
                content = file.read()
                speak("Showing Notes")
                ui_queue.put(("system_output", "Showing Notes"))
                print(content)
                speak(content)
                ui_queue.put(("system_output", content))
        except Exception as e:
            speak("I couldn't find any notes.")
            ui_queue.put(("system_output", "I couldn't find any notes."))

    elif "update assistant" in query:
        try:
            speak("After downloading the file, please replace this file with the downloaded one.")
            ui_queue.put(
                ("system_output", "After downloading the file, please replace this file with the downloaded one."))
            # Provide the actual URL after uploading the file
            url = 'https://your-update-url.com/Voice.py'
            response = requests.get(url, stream=True)
            with open("Voice.py", "wb") as file:
                total_length = int(response.headers.get('content-length', 0))
                for data in response.iter_content(chunk_size=4096):
                    file.write(data)
            speak("Assistant updated successfully.")
            ui_queue.put(("system_output", "Assistant updated successfully."))
        except Exception as e:
            speak("I couldn't update the assistant.")
            ui_queue.put(("system_output", "I couldn't update the assistant."))

    elif "misty" in query:
        speak("Misty at your service.")
        ui_queue.put(("system_output", "Misty at your service."))

    elif "weather" in query:
        try:
            api_key = "YOUR_OPENWEATHERMAP_API_KEY"
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak("City name")
            ui_queue.put(("system_output", "City name"))
            city_name = take_command()
            complete_url = f"{base_url}appid={api_key}&q={city_name}"
            response = requests.get(complete_url)
            x = response.json()

            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidity = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                weather_info = (f"Temperature (in kelvin unit) = {current_temperature}\n"
                                f"Atmospheric pressure (in hPa unit) = {current_pressure}\n"
                                f"Humidity (in percentage) = {current_humidity}\n"
                                f"Description = {weather_description}")
                print(weather_info)
                speak(weather_info)
                ui_queue.put(("system_output", weather_info))
            else:
                speak("City Not Found")
                ui_queue.put(("system_output", "City Not Found"))
        except Exception as e:
            speak("I couldn't fetch the weather information.")
            ui_queue.put(("system_output", "I couldn't fetch the weather information."))

    elif "send message" in query:
        try:
            # You need to set up Twilio and replace the placeholders
            from twilio.rest import Client
            account_sid = 'YOUR_TWILIO_ACCOUNT_SID'
            auth_token = 'YOUR_TWILIO_AUTH_TOKEN'
            client = Client(account_sid, auth_token)

            speak("What should the message say?")
            ui_queue.put(("system_output", "What should the message say?"))
            message_body = take_command()
            message = client.messages.create(
                body=message_body,
                from_='YOUR_TWILIO_PHONE_NUMBER',
                to='RECIPIENT_PHONE_NUMBER'
            )
            print(message.sid)
            speak("Message has been sent.")
            ui_queue.put(("system_output", "Message has been sent."))
        except Exception as e:
            speak("I couldn't send the message.")
            ui_queue.put(("system_output", "I couldn't send the message."))

    elif "good morning" in query:
        speak("A warm Good Morning to you!")
        ui_queue.put(("system_output", "A warm Good Morning to you!"))

    elif "will you be my gf" in query or "will you be my bf" in query:
        speak("I'm not sure about, maybe you should give me some time.")
        ui_queue.put(("system_output", "I'm not sure about, maybe you should give me some time."))

    elif "i love you" in query:
        speak("It's hard to understand. But, I love to help you dear!")
        ui_queue.put(("system_output", "It's hard to understand."))

    elif "Bad words" in query:
        speak("Madarchod bokachoda khankir chele, gud marani pod marani rendir chele,"
              " bichir bal guder chal dhoyachoda banchod,"
              " beshamagi kutta choda hati choda dinosaur choda,"
              " angel choda virus choda ut khankir dim,"
              " panty chor condom chor gudir chele classical choda, mobile choda chuyar choda tepon magi rendi magir gud chete fel dudh")
        ui_queue.put(("system_output", "It's hard to understand."))

    elif "terminate mouse" in query:
        terminate_process('py.exe')

    elif "close microsoft word" in query:
        terminate_process('WINWORD.EXE')

    elif "close active window" in query:
        try:
            pyautogui.hotkey('alt', 'f4')
            speak("Closed the active window.")
            ui_queue.put(("system_output", "Closed the active window."))
        except Exception as e:
            speak("I couldn't close the active window.")
            ui_queue.put(("system_output", "I couldn't close the active window."))

    elif "open" in query:
        try:
            app_name = query.replace("open", "").strip()
            speak(f"Opening {app_name}")
            ui_queue.put(("system_output", f"Opening {app_name}"))

            pyautogui.press('win')
            pyautogui.press('space')
            query = query.replace("open", "")
            pyautogui.write(app_name)
            time.sleep(2)
            pyautogui.press('enter')
        except Exception as e:
            speak(f"Unable to open {app_name}.")
            ui_queue.put(("system_output", f"Unable to open {app_name}."))

    elif "write" in query or "right" in query or "type" in query:
        try:
            text = query.replace("write", "").replace("right", "").replace("type", "").strip()
            pyautogui.write(text, interval=0.05)
            speak(f"Written {text}")
            ui_queue.put(("system_output", f"Written {text}"))
        except Exception as e:
            speak("I couldn't type that.")
            ui_queue.put(("system_output", "I couldn't type that."))


## keyboard Control-----------------------------------------------
    elif "enter" in query or "inter" in query:
        try:
            pyautogui.press('enter')
            speak("Pressed Enter.")
            ui_queue.put(("system_output", "Pressed Enter."))
        except Exception as e:
            speak("I couldn't press Enter.")
            ui_queue.put(("system_output", "I couldn't press Enter."))

    elif "go back" in query:
        try:
            pyautogui.press('backspace')
            speak("Went back.")
            ui_queue.put(("system_output", "Went back."))
        except Exception as e:
            speak("I couldn't go back.")
            ui_queue.put(("system_output", "I couldn't go back."))

    elif "go down" in query:
        try:
            pyautogui.press('down')
            speak("Went down.")
            ui_queue.put(("system_output", "Went down."))
        except Exception as e:
            speak("I couldn't go down.")
            ui_queue.put(("system_output", "I couldn't go down."))

    elif "go up" in query:
        try:
            pyautogui.press('up')
            speak("Went up.")
            ui_queue.put(("system_output", "Went up."))
        except Exception as e:
            speak("I couldn't go up.")
            ui_queue.put(("system_output", "I couldn't go up."))

    elif "select" in query:
        try:
            pyautogui.hotkey('ctrl', 'a')
            speak("Selected all.")
            ui_queue.put(("system_output", "Selected all."))
        except Exception as e:
            speak("I couldn't select.")
            ui_queue.put(("system_output", "I couldn't select."))

    elif "erase" in query:
        try:
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            speak("Erased the content.")
            ui_queue.put(("system_output", "Erased the content."))
        except Exception as e:
            speak("I couldn't erase.")
            ui_queue.put(("system_output", "I couldn't erase."))

    elif "copy" in query:
        try:
            pyautogui.hotkey('ctrl', 'c')
            speak("Copied.")
            ui_queue.put(("system_output", "Copied."))
        except Exception as e:
            speak("I couldn't Copy.")
            ui_queue.put(("system_output", "I couldn't Copy."))

    elif "paste" in query:
        try:
            pyautogui.hotkey('ctrl', 'v')
            speak("Pasted.")
            ui_queue.put(("system_output", "Pasted."))
        except Exception as e:
            speak("I couldn't Copy.")
            ui_queue.put(("system_output", "I couldn't Copy."))

    # Mouse Controls
    elif "left click" in query:
        try:
            pyautogui.click(button='left')
            speak("Left click done.")
            ui_queue.put(("system_output", "Left click done."))
        except Exception as e:
            speak("I couldn't perform left click.")
            ui_queue.put(("system_output", "I couldn't perform left click."))

    elif "right click" in query:
        try:
            pyautogui.click(button='right')
            speak("Right click done.")
            ui_queue.put(("system_output", "Right click done."))
        except Exception as e:
            speak("I couldn't perform right click.")
            ui_queue.put(("system_output", "I couldn't perform right click."))

    elif "double click" in query:
        try:
            pyautogui.doubleClick()
            speak("Double click done.")
            ui_queue.put(("system_output", "Double click done."))
        except Exception as e:
            speak("I couldn't perform double click.")
            ui_queue.put(("system_output", "I couldn't perform double click."))


    #-------------------------------Cordinates X Y Movement----------------------------------------------------------------------
    elif "deactivate screen mouse" in query or "deactivate grid screen" in query:
        threading.Thread(target=close_grid_screen).start()
        speak("Grid Screen X, Y Position View Closed")
        ui_queue.put(("system_output", "Grid Screen X, Y Position View Closed"))

    elif "activate screen mouse" in query or "activate grid screen" in query:
        # Start a new thread to open the screen mouse window
        threading.Thread(target=open_grid_screen).start()
        speak("Grid Screen X, Y Position View Activated")
        ui_queue.put(("system_output", "Grid Screen X, Y Position View opened"))
        # Check if the user said "close screen mouse"


    elif "move mouse to" in query or "mouse to" in query or "mouse" in query:
        try:
            # Extract the part of the query after "move mouse to"
            coords_part = query.replace("move mouse to", "").strip()

            # Replace any non-numeric, non-space characters with a space
            coords_part = re.sub(r'[^\d\s]', ' ', coords_part)

            # Split the cleaned-up part into components
            coords = coords_part.split()

            if len(coords) == 2:
                try:
                    # Convert the coordinates to integers
                    x, y = int(coords[0]), int(coords[1])

                    # Move the mouse to the specified coordinates
                    pyautogui.moveTo(x, y, duration=0.5)
                    speak(f"Moved mouse to {x}, {y}")
                    ui_queue.put(("system_output", f"Moved mouse to {x}, {y}"))

                except ValueError:
                    # Handle cases where coordinates are not valid integers
                    speak("The coordinates should be numbers. Please try again.")
                    ui_queue.put(("system_output", "The coordinates should be numbers. Please try again."))

            else:
                # Handle incorrect number of coordinates
                speak("Please provide exactly two coordinates separated by a space.")
                ui_queue.put(("system_output", "Please provide exactly two coordinates separated by a space."))

        except Exception as e:
            # General exception handling with error details
            speak("I couldn't move the mouse due to an error.")
            ui_queue.put(("system_output", "I couldn't move the mouse due to an error."))
            # Optionally, log the exception
            print(f"Error: {e}")


    ### ------------------------------------Commands For Mouse. Goes from here.------------------------------------------
    # left click
    elif "left click" in query:
        mouse.click('left')
        speak("Done")

        # right click
    elif "right click" in query:
        mouse.click('right')
        speak("Done")

    # middle click
    elif "middle click" in query:
        mouse.click('middle')
        speak("Done")
    # double click
    elif "double click" in query:
        mouse.double_click()
        speak("Done")
    # get the position of mouse
    elif "position" in query:
        print(mouse.get_position())
        speak(mouse.get_position())

    # In [12]: mouse.get_position()
    # Out[12]: (714, 488)

    # presses but doesn't release
    elif "press and hold" in query:
        mouse.hold('left')
        speak("Done")

    elif "hold release" in query:
        mouse.release('left')
        speak("Done")
    # mouse.press('left')

    # drag from (0, 0) to (100, 100) relatively with a duration of 0.1s
    elif "drag" in query or "drug" in query:
        mouse.drag(0, 0, 100, 100, absolute=False, duration=0.1)
        speak("Done")

    # whether a button is clicked
    # print(mouse.is_pressed('right'))

    # move 100 right & 100 down..................................................................
    # move Right ..................................................
    elif "move top right" in query:
        mouse.move(1896, 17, absolute=True, duration=0.1)
        speak("Done")
        #continue
    elif "move bottom right" in query:
        mouse.move(1888, 1045, absolute=True, duration=0.1)
        speak("Done")
        #continue
    elif "move right down" in query:
        mouse.move(100, 100, absolute=False, duration=0.1)
        speak("Done")
        #continue
    elif "move right side up" in query:
        mouse.move(100, -100, absolute=False, duration=0.1)
        speak("Done")
        #continue

    elif "move right" in query:
        mouse.move(100, 0, absolute=False, duration=0.1)
        speak("Done")
 #       continue


    # move Left .............................................................................
    elif "move top left" in query:
        mouse.move(33, 20, absolute=True, duration=0.1)
        speak("Done")
#        continue
    elif "move bottom left" in query:
        mouse.move(27, 1056, absolute=True, duration=0.1)
        speak("Done")
 #       continue
    elif "move left up" in query:
        mouse.move(-100, -100, absolute=False, duration=0.1)
        speak("Done")
  #      continue
    elif "move left down" in query:
        mouse.move(-100, 100, absolute=False, duration=0.1)
        speak("Done")
    #    continue
    elif "move left" in query:
        mouse.move(-100, 0, absolute=False, duration=0.3)
        speak("Done")
   #     continue


    # move up and down.....................
    elif "move up" in query:
        mouse.move(0, -100, absolute=False, duration=0.1)
        speak("Done")
     #   continue
    elif "move down" in query:
        mouse.move(0, 100, absolute=False, duration=0.1)
        speak("Done")
      #  continue

    # move little........................................................
    elif "little bit left" in query:
        mouse.move(-22, 0, absolute=False, duration=0.1)
        speak("Done")
#        continue
    elif "little bit right" in query:
        mouse.move(22, 0, absolute=False, duration=0.1)
        speak("Done")
 #       continue
    elif "little bit up" in query:
        mouse.move(0, -22, absolute=False, duration=0.1)
        speak("Done")
  #      continue
    elif "little bit down" in query:
        mouse.move(0, 22, absolute=False, duration=0.1)
        speak("Done")
   #     continue

    elif "little left" in query:
        mouse.move(-45, 0, absolute=False, duration=0.1)
        speak("Done")
#        continue
    elif "little right" in query:
        mouse.move(45, 0, absolute=False, duration=0.1)
        speak("Done")
 #       continue
    elif "little top" in query or "little up" in query or "little app" in query:
        mouse.move(0, -45, absolute=False, duration=0.1)
        speak("Done")
  #      continue
    elif "little down" in query:
        mouse.move(0, 45, absolute=False, duration=0.1)
        speak("Done")
   #     continue

    # Move more extra ........................................................
    elif "double left" in query:
        mouse.move(-500, 0, absolute=False, duration=0.1)
        speak("Done")
    #    continue
    elif "double right" in query:
        mouse.move(500, 0, absolute=False, duration=0.1)
        speak("Done")
     #   continue
    elif "double up" in query:
        mouse.move(0, -500, absolute=False, duration=0.1)
        speak("Done")
      #  continue
    elif "double down" in query:
        mouse.move(0, 500, absolute=False, duration=0.1)
        speak("Done")
       # continue
    # ....................................Extra move.................
    elif "extra left" in query:
        mouse.move(-245, 0, absolute=False, duration=0.1)
        speak("Done")
        #continue
    elif "extra right" in query:
        mouse.move(245, 0, absolute=False, duration=0.1)
        speak("Done")
        #continue
    elif "extra up" in query:
        mouse.move(0, -245, absolute=False, duration=0.1)
        speak("Done")
        #continue
    elif "extra down" in query:
        mouse.move(0, 245, absolute=False, duration=0.1)
        speak("Done")
        #continue



    # move middle ...............................................................
    elif "move middle top" in query:
        mouse.move(945, 22, absolute=True, duration=0.1)
        speak("Done")
        #continue
    elif "move bottom middle" in query:
        mouse.move(962, 1054, absolute=True, duration=0.1)
        speak("Done")
        #continue
    elif "move middle" in query:
        mouse.move(961, 458, absolute=True, duration=0.1)
        speak("Done")
        #continue

    # make a listener when left button is clicked
    # mouse.on_click(lambda: print("Left Button clicked."))
    # make a listener when right button is clicked
    # mouse.on_right_click(lambda: print("Right Button clicked."))

    # remove the listeners when you want
    # mouse.unhook_all()

    # scroll down
    elif "long down scroll" in query:
        mouse.wheel(-3)
        speak("ok")
#        continue
    # scroll up
    elif "long up scroll" in query:
        mouse.wheel(3)
        speak("Done")
 #       continue
    elif "scroll down" in query:
        mouse.wheel(-1)
        speak("ok")
  #      continue
    elif "scroll up" in query:
        mouse.wheel(1)
        speak("ok")
   #     continue

    #----------------------------------------Some Flexibility Commands--------------
    elif 'zoom in' in query:
        pyautogui.hotkey('ctrl', '+')
        ui_queue.put("Zooming in.")

    elif 'zoom out' in query:
        pyautogui.hotkey('ctrl', '-')
        ui_queue.put("Zooming out.")

    elif 'minimize' in query or 'minimise' in query:
        pyautogui.hotkey('win', 'down')
        ui_queue.put("Minimizing the tab.")

    elif 'maximize' in query or 'maximise' in query:
        pyautogui.hotkey('win', 'up')
        ui_queue.put("Maximizing the tab.")

    elif 'switch tab' in query:
        pyautogui.hotkey('alt', 'tab')
        ui_queue.put("Switching between tabs.")

        # Volume up or down to a specific value

    elif 'volume up to' in query or 'volume down to' in query:

        volume_level = get_level_from_query(query)

        set_volume(volume_level)

        ui_queue.put(f"Setting volume to {volume_level}%.")

        speak(f"Setting volume to {volume_level}%.")


    elif 'brightness up to' in query or 'brightness down to' in query:

        brightness_level = get_level_from_query(query)

        sbc.set_brightness(brightness_level)

        ui_queue.put(f"Setting brightness to {brightness_level}%.")

        speak(f"Setting brightness to {brightness_level}%.")


    elif 'bluetooth on' in query:
        subprocess.call('powershell.exe Start-Service bthserv', shell=True)
        ui_queue.put("Turning on Bluetooth.")

    elif 'bluetooth off' in query:
        subprocess.call('powershell.exe Stop-Service bthserv', shell=True)
        ui_queue.put("Turning off Bluetooth.")

    elif 'wi-fi on' in query:
        subprocess.call('netsh interface set interface "Wi-Fi" enabled', shell=True)
        ui_queue.put("Turning on Wi-Fi.")

    elif 'wi-fi off' in query:
        subprocess.call('netsh interface set interface "Wi-Fi" disabled', shell=True)
        ui_queue.put("Turning off Wi-Fi.")

    elif 'open camera' in query:
        subprocess.call('start microsoft.windows.camera:', shell=True)
        ui_queue.put("Opening the camera.")

    elif 'close camera' in query:
        subprocess.call('taskkill /IM WindowsCamera.exe /F', shell=True)
        ui_queue.put("Closing the camera.")

    elif 'mute mic' in query:
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMute(1, None)
        ui_queue.put("Muting the microphone.")

    elif 'unmute mic' in query:
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMute(0, None)
        ui_queue.put("Unmuting the microphone.")

    # record until you click right

    # events = mouse.record()

    # replay these events
    # mouse.play(events[:-1])

    # elif "" in query:
    # Command go here
    # For adding more commands

    else:
        speak("Sorry, I didn't understand that.")
        ui_queue.put(("system_output", "Sorry, I didn't understand that."))



# Initialize Bot Name
bot_name = "Misty"


# ----------------------- UI Class -----------------------

class AssistantUI:
    """

    Class representing the user interface for the Misty Assistant application.

    class AssistantUI:
        def __init__(self, root, ui_queue):
            """
    def __init__(self, root, ui_queue):
        self.root = root
        self.ui_queue = ui_queue

        self.root.title("Misty Assistant")
        self.root.geometry("400x500")
        self.root.configure(bg='#2C3E50')  # Dark background
        self.root.resizable(False, False)

        # Load Icons
        try:
            self.listen_icon = PhotoImage(file=LISTEN_ICON_PATH)
            self.recognize_icon = PhotoImage(file=RECOGNIZE_ICON_PATH)
        except Exception as e:
            print("Icon files not found. Please ensure listen_icon.png and recognize_icon.png are in the directory.")
            self.listen_icon = None
            self.recognize_icon = None

        # Status Frame
        self.status_frame = tk.Frame(self.root, bg='#2C3E50')
        self.status_frame.pack(pady=10)

        # Listening Indicator
        if self.listen_icon:
            self.listen_label = tk.Label(self.status_frame, image=self.listen_icon, bg='#2C3E50')
            self.listen_label.grid(row=0, column=0, padx=5)
        else:
            self.listen_label = tk.Label(self.status_frame, text="Listening", fg='white', bg='#2C3E50')
            self.listen_label.grid(row=0, column=0, padx=5)

        # Recognizing Indicator
        if self.recognize_icon:
            self.recognize_label = tk.Label(self.status_frame, image=self.recognize_icon, bg='#2C3E50')
            self.recognize_label.grid(row=0, column=1, padx=5)
        else:
            self.recognize_label = tk.Label(self.status_frame, text="Recognizing", fg='white', bg='#2C3E50')
            self.recognize_label.grid(row=0, column=1, padx=5)

        # User Input Display
        self.user_input_label = tk.Label(self.root, text="User Input:", bg='#2C3E50', fg='white',
                                         font=("Arial", 12, "bold"))
        self.user_input_label.pack(anchor='w', padx=20)

        self.user_input_box = ScrolledText(self.root, wrap=tk.WORD, height=8, font=("Arial", 10), state='disabled',
                                           bg='#34495E', fg='white')
        self.user_input_box.pack(pady=5, padx=20, fill='both')

        # System Output Display
        self.system_output_label = tk.Label(self.root, text="System Output:", bg='#2C3E50', fg='white',
                                            font=("Arial", 12, "bold"))
        self.system_output_label.pack(anchor='w', padx=20)

        self.system_output_box = ScrolledText(self.root, wrap=tk.WORD, height=8, font=("Arial", 10), state='disabled',
                                              bg='#34495E', fg='white')
        self.system_output_box.pack(pady=5, padx=20, fill='both')

        # Initialize UI Update
        self.update_ui()

    def update_ui(self):
        """Method to update the UI based on the queue messages."""
        try:
            while True:
                message = self.ui_queue.get_nowait()
                if message[0] == "status":
                    status = message[1]
                    if status == "listening":
                        self.set_listening(True)
                        self.set_recognizing(False)
                    elif status == "recognizing":
                        self.set_listening(False)
                        self.set_recognizing(True)
                    else:
                        self.set_listening(False)
                        self.set_recognizing(False)
                elif message[0] == "user_input":
                    self.append_text(self.user_input_box, f"User: {message[1]}")
                elif message[0] == "system_output":
                    self.append_text(self.system_output_box, f"Misty: {message[1]}")
        except queue.Empty:
            pass
        self.root.after(100, self.update_ui)

    def set_listening(self, is_listening):
        """Update the listening indicator."""
        if self.listen_icon:
            if is_listening:
                self.listen_label.config(bg='green')
            else:
                self.listen_label.config(bg='#2C3E50')
        else:
            if is_listening:
                self.listen_label.config(fg='green')
            else:
                self.listen_label.config(fg='white')

    def set_recognizing(self, is_recognizing):
        """Update the recognizing indicator."""
        if self.recognize_icon:
            if is_recognizing:
                self.recognize_label.config(bg='green')
            else:
                self.recognize_label.config(bg='#2C3E50')
        else:
            if is_recognizing:
                self.recognize_label.config(fg='green')
            else:
                self.recognize_label.config(fg='white')

    def append_text(self, text_widget, text):
        """Append text to a ScrolledText widget."""
        text_widget.config(state='normal')
        text_widget.insert(tk.END, text + "\n")
        text_widget.config(state='disabled')
        text_widget.see(tk.END)


# ----------------------- Main Function -----------------------

def main():
    """
    Initializes the main components of the application GUI, sets up the assistant thread,
    cleans any command before executing the python file, and starts the main event loop.

    :return: None
    """
    global ui_queue
    ui_queue = queue.Queue()

    root = tk.Tk()
    app = AssistantUI(root, ui_queue)

    # Start the assistant thread
    assistant_thread = threading.Thread(target=take_command, args=(queue.Queue(), ui_queue), daemon=True)
    assistant_thread.start()


    # This Function will clean any
    # command before execution of this python file
    clear: Callable[[], int] = lambda: os.system('cls')
    clear()
    wishme()
    #username()

    root.mainloop()


if __name__ == "__main__":
    main()
