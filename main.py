import openai
import pyttsx3
import os
import subprocess
import datetime
import speech_recognition as sr
import webbrowser
import pyautogui
import requests
import pywhatkit
import time
import keyboard
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from config import apikey

chatStr = ""
memory = {}


def initialize_text_to_speech():
    engine_instance = pyttsx3.init()
    return engine_instance


def say(text, speech_engine):
    speech_engine.say(text)
    speech_engine.runAndWait()


def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    try:
        response = openai.Completion.create(
            model="gpt-3.5-turbo-instruct",
            prompt=prompt,
            temperature=1,
            max_tokens=256,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )

        text += response["choices"][0]["text"]
        if not os.path.exists("Openai"):
            os.mkdir("Openai")

        with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
            f.write(text)

    except Exception as e:
        print(f"Error in OpenAI API call: {str(e)}")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            audio = r.listen(source)
            print("Recognizing...")
            user_query = r.recognize_google(audio, language="en-in")
            print(f"User Said: {user_query}")

            if "jarvis stop" in user_query.lower():
                print("User interrupted Jarvis.")
                return None

            return user_query.lower()

        except KeyboardInterrupt:
            print("Program interrupted by user.")
            return None

        except sr.UnknownValueError:
            print("Speech Recognition could not understand the audio")
            return None

        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {str(e)}")
            return None

        except Exception as e:
            print(f"An unexpected error occurred: {str(e)}")
            return None


def stop_jarvis_talking():
    global listening_state
    listening_state = True
    if speech_engine.isBusy():
        speech_engine.stop()


keyboard.add_hotkey('ctrl+alt', stop_jarvis_talking)


def set_volume(volume_level):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(volume_level, None)


set_volume(0.5)


def type_text(user_query):
    pyautogui.typewrite(user_query)


def press_key(key):
    pyautogui.press(key)


def go_back():
    pyautogui.hotkey('alt', 'left')


def go_down():
    pyautogui.hotkey('down')


def go_up():
    pyautogui.hotkey('up')


def go_next():
    pyautogui.hotkey('alt', 'right')


def close_tab():
    pyautogui.hotkey('ctrl', 'w')


def backspace():
    pyautogui.press('backspace')


def run_code():
    pyautogui.hotkey('fn', 'shift', 'F10')


def vs_code():
    pyautogui.hotkey('fn', 'F5')


def shift():
    pyautogui.hotkey('shift')


def click_mouse():
    pyautogui.click()


def right_click_mouse():
    pyautogui.rightClick()


def double_click_mouse():
    pyautogui.doubleClick()


def scroll_up():
    pyautogui.scroll(1)


def scroll_down():
    pyautogui.scroll(-1)


def select_all():
    pyautogui.hotkey('ctrl', 'a')


def copy():
    pyautogui.hotkey('ctrl', 'c')


def cut():
    pyautogui.hotkey('ctrl', 'x')


def paste():
    pyautogui.hotkey('ctrl', 'v')


def new_file():
    pyautogui.hotkey('ctrl', 'n')


def save():
    pyautogui.hotkey('ctrl', 's')


def play_music(query):
    try:
        pywhatkit.playonyt(query)
        return f"Playing {query}"
    except Exception as e:
        return f"Error playing music: {str(e)}"


def close_current_tab():
    pyautogui.hotkey('fn', 'alt', 'f4')
    say('Closing the tab, Sir', speech_engine)


def type_special_character(character):
    pyautogui.typewrite(character)
    print(f"Typed {character}")


def get_weather():
    try:
        api_key = 'adce0d78c3f7c01d45fb66e78f1d2f7b'
        city = 'Pune, IN'

        url = f'https://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}'
        response = requests.get(url)
        data = response.json()

        if data['cod'] == '404':
            return "City not found. Please check the city name."

        temperature = data['main']['temp']
        description = data['weather'][0]['description']

        return f"The current temperature in {city} is {temperature}°C with {description}."

    except Exception as e:
        return f"Error retrieving weather information: {str(e)}"


def chat(user_query, speech_engine):
    global chatStr
    global memory
    openai.api_key = apikey
    chatStr += f"User: {user_query}\n Jarvis: "

    if "remember" in user_query:
        info_to_remember = user_query.split("remember")[-1].strip()

        memory["remembered_info"] = info_to_remember

        reply_text = f"I will remember: {info_to_remember}"
        say(reply_text, speech_engine)
        chatStr += f"{reply_text}\n"
        return reply_text

    elif "recall" in user_query:
        remembered_info = memory.get("remembered_info", "I don't remember anything.")
        say(remembered_info, speech_engine)
        chatStr += f"{remembered_info}\n"
        return remembered_info

    elif "how are you" in user_query:
        reply_text = "I'm doing well, thank you for asking! How can I help you today?"
        say(reply_text, speech_engine)
        chatStr += f"{reply_text}\n"
        return reply_text

    elif "how you are able to generate code" in user_query:
        reply_text = "I am ai model trained to generate code"
        say(reply_text, speech_engine)
        chatStr += f"{reply_text}\n"
        return reply_text

    elif "weather" in user_query:
        weather_info = get_weather()
        say(weather_info, speech_engine)
        chatStr += f"{weather_info}\n"
        return weather_info

    else:
        try:
            response = openai.Completion.create(
                model="gpt-3.5-turbo-instruct",
                prompt=user_query,
                temperature=0.8,
                max_tokens=256,
                top_p=1,
                frequency_penalty=0,
                presence_penalty=0
            )

            reply_text = response["choices"][0]["text"]
            say(reply_text, speech_engine)
            remembered_info = memory.get("remembered_info", None)
            if remembered_info is not None:
                say(f"I also want to remind you: {remembered_info}", speech_engine)
                chatStr += f"I also want to remind you: {remembered_info}\n"
                memory["remembered_info"] = None

            chatStr += f"{reply_text}\n"
            return reply_text

        except Exception as e:
            print(f"Error in OpenAI API call: {str(e)}")
            return "I'm sorry, but I couldn't generate a response at the moment."


def get_locations(api_key, location_query):
    try:
        url = f'https://maps.googleapis.com/maps/api/place/textsearch/json?query={location_query}&key={api_key}'
        response = requests.get(url)
        data = response.json()

        locations = []
        for result in data['results']:
            place_name = result['name']
            address = result.get('formatted_address', 'Address not available')
            locations.append(f"{place_name} at {address}")

        if not locations:
            return f"No locations found for '{location_query}'."

        return f"Here are some locations matching '{location_query}':\n" + '\n'.join(locations)

    except Exception as e:
        return f"Error retrieving location information: {str(e)}"


def open_folder(folder_path):
    try:
        os.startfile(folder_path)
        say(f"Opening folder: {folder_path}", speech_engine)
    except Exception as e:
        say(f"Error opening folder: {str(e)}", speech_engine)
        folder_path_to_open = r"C:\Users\91952\Desktop\Openai"
        open_folder(folder_path_to_open)


def open_whatsapp():
    webbrowser.open("https://web.whatsapp.com/")


def send_whatsapp_message(contact_name, message):
    try:
        pywhatkit.sendwhatmsg_instantly(contact_name, message, wait_time=10)
        return f"Message sent to {contact_name}: {message}"
    except Exception as e:
        return f"Error sending WhatsApp message: {str(e)}"


if __name__ == '__main__':
    speech_engine = initialize_text_to_speech()
    listening_state = False

    while True:
        print("Listening...")
        user_query = takeCommand()

        if user_query is None:
            print("Please repeat your query.")
            continue

        if "hello jarvis" in user_query:
            listening_state = True
            print("Listening activated.")
            response_text = "Karan, welcome to Jarvis A.I. How can I help you today?"
            say(response_text, speech_engine)
            continue

        if "jarvis top" in user_query:
            listening_state = False
            print("jarvis is stopped.")
            say("jarvis is stopped.", speech_engine)
            continue

        elif "jarvis quit" in user_query:
            listening_state = False
            print("jarvis is quitting. Goodbye!")
            say("jarvis is quitting. Goodbye!", speech_engine)
            exit()

        elif "jarvis wake up" in user_query:
            listening_state = True
            print("Karan, welcome to Jarvis A.I. How can I help you today?")
            say("Karan, welcome to Jarvis A.I. How can I help you today?", speech_engine)
            continue

        if listening_state:
            if "open calculator" in user_query:
                subprocess.Popen("calc.exe", shell=True)
                say('Opening Calculator, Sir', speech_engine)

            elif "open vs code" in user_query:
                vscode_path = r"C:\\Users\\91952\\Desktop\\Visual Studio Code.lnk"
                subprocess.Popen(vscode_path, shell=True)
                say('Opening Visual Studio Code Sir', speech_engine)

            elif "open pycharm" in user_query:
                pycharm_path = r"C:\Users\Public\Desktop\PyCharm Community Edition 2023.3.3.lnk"
                subprocess.Popen(pycharm_path, shell=True)
                say('Opening Pycharm Sir', speech_engine)

            elif "open android studio" in user_query:
                android_path = r"C:\Users\91952\Desktop\Android Studio.lnk"
                subprocess.Popen(android_path, shell=True)
                say('Opening Android Studio Sir', speech_engine)

            elif "open asphalt 9" in user_query:
                asphalt_path = r"C:\Users\91952\Desktop\Asphalt 9 Legends - Shortcut.lnk"
                subprocess.Popen(asphalt_path, shell=True)
                say('Opening Asphalt 9 Sir', speech_engine)

            elif "open notepad" in user_query:
                notepad_path = r"C:\Users\Public\Desktop\Notepad++.lnk"
                subprocess.Popen(notepad_path, shell=True)
                say('Opening Notepad Sir', speech_engine)

            elif "open chrome" in user_query:
                subprocess.Popen("start chrome", shell=True)
                say('Opening Chrome Sir', speech_engine)

            elif "open youtube" in user_query:
                say("Search for:", speech_engine)
                search = takeCommand()
                if search is None:
                    webbrowser.open("https://www.youtube.com/")
                    say("I'm not able to get your search command.", speech_engine)
                    say("I'm opening the YouTube homepage.", speech_engine)
                    say("Opening YouTube.", speech_engine)
                else:
                    webbrowser.open("https://www.youtube.com/results?search_query=" + search)
                    say('Opening search results on YouTube.', speech_engine)

            elif "open gmail" in user_query:
                webbrowser.open("https://mail.google.com/")
                say('Opening Gmail Sir', speech_engine)

            elif "open command prompt" in user_query:
                subprocess.Popen("cmd.exe", shell=True)
                say('Opening Command Prompt Sir', speech_engine)

            elif "search on google" in user_query:
                query_terms = user_query.replace("search on google", "").strip()
                webbrowser.open(f"https://www.google.com/search?q={query_terms}")

            elif "using artificial intelligence" in user_query:
                ai(prompt=user_query)

            elif "jarvis close" in user_query:
                close_current_tab()

            elif "jarvis shut down pc" in user_query:
                close_current_tab()

            elif "type" in user_query:
                text_to_type = user_query.replace("type", "").strip()
                type_text(text_to_type)

            elif "press enter" in user_query:
                press_key("enter")

            elif "go back" in user_query:
                go_back()

            elif "go next" in user_query:
                go_next()

            elif "go up" in user_query:
                go_up()

            elif "go down" in user_query:
                go_down()

            elif "backspace" in user_query:
                backspace()

            elif "jarvis give hash" in user_query:
                type_special_character('#')

            elif "jarvis give semicolon" in user_query:
                type_special_character(';')

            elif "jarvis give parenthesis" in user_query:
                type_special_character('(')

            elif "jarvis give curly brace" in user_query:
                type_special_character('{')

            elif "jarvis give square bracket" in user_query:
                type_special_character('[')

            elif "jarvis give double coat" in user_query:
                type_special_character('"')

            elif "jarvis play" in user_query:
                music_query = user_query.replace("jarvis play", "").strip()
                response = play_music(music_query)
                say(response, speech_engine)

            elif "click mouse" in user_query:
                click_mouse()
                say("Mouse clicked", speech_engine)

            elif "right click mouse" in user_query:
                right_click_mouse()
                say("Right-click performed", speech_engine)

            elif "double click mouse" in user_query:
                double_click_mouse()
                say("Double-click performed", speech_engine)

            elif "scroll up" in user_query:
                scroll_up()
                say("Mouse scrolled up", speech_engine)

            elif "scroll down" in user_query:
                scroll_down()
                say("Mouse scrolled down", speech_engine)

            elif "select all" in user_query:
                select_all()
                say("Text selected", speech_engine)

            elif "select on" in user_query:
                say("Text selected", speech_engine)

            elif "copy" in user_query:
                copy()
                say("Text copied", speech_engine)

            elif "cut" in user_query:
                cut()
                say("Text cut", speech_engine)

            elif "drop" in user_query:
                paste()
                say("Text pasted", speech_engine)

            elif "new file" in user_query:
                new_file()
                say("New file created", speech_engine)

            elif "shift" in user_query:
                shift()
                say("Shift pressed", speech_engine)

            elif "save" in user_query:
                save()
                say("File saved", speech_engine)

            elif "open whatsapp" in user_query:
                open_whatsapp()
                say('Opening Whatsapp Sir', speech_engine)

            elif "send message on whatsapp" in user_query:
                contact_name = "+919529491387"
                message = user_query.split("send message")[-1].strip()
                response = send_whatsapp_message(contact_name, message)
                say("Message send", speech_engine)

            elif "open folder" in user_query:
                folder_path = user_query.replace("open folder", "").strip()
                open_folder(folder_path)
                say("Folder opened", speech_engine)

            elif "set volume to" in user_query:
                try:
                    volume_level = float(user_query.split("set volume to")[-1].strip()) / 100
                    set_volume(volume_level)
                    say(f"Volume set to {int(volume_level * 100)}%", speech_engine)
                except ValueError:
                    say("Invalid volume level. Please specify a number between 0 and 100.", speech_engine)

            elif "jarvis stop" in user_query:
                say("I'm not able to get your command.", speech_engine)

            elif "the time" in user_query:
                hour = datetime.datetime.now().strftime("%H")
                minute = datetime.datetime.now().strftime("%M")
                say(f"Sir, the time is {hour} hours and {minute} minutes.", speech_engine)

            elif "weather" in user_query:
                weather_info = get_weather()
                say(weather_info, speech_engine)
                chatStr += f"{weather_info}\n"

            elif "locations" in user_query:
                location_query = user_query.replace("jarvis, tell me about", "").strip()
                api_key = 'GOOGLE_MAPS_API_KEY'
                response = get_locations(api_key, location_query)
                say(response, speech_engine)

            elif "reset chat" in user_query:
                chatStr = ""

            elif "save chat history" in user_query:
                with open("chat_history.txt", "w") as file:
                    file.write(chatStr)
                say('Chat history saved, Sir', speech_engine)

            elif "write a program" in user_query and "in" in user_query and "for" in user_query:
                language = user_query.split("in")[1].split("for")[0].strip()
                task = user_query.split("for")[1].strip()
                say(f"Sure! Please provide more details for the {language} program {task}.", speech_engine)
                user_program = takeCommand()
                if user_program:
                    prompt = f"Write a {language} program for {task}: {user_program}"
                    response = chat(prompt, speech_engine)
                    if response:
                        print("Generated Program:")
                        say(f"Here's the {language} program I generated for {task}:", speech_engine)
                        time.sleep(2)
                        pyautogui.typewrite(response, interval=0.02)
                    else:
                        print("Failed to generate program.")

            elif "jarvis run" in user_query:
                run_code()

            elif "run this" in user_query:
                vs_code()

            else:
                response = chat(user_query, speech_engine)