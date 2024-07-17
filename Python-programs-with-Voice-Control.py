import subprocess
import os
import shlex
import speech_recognition as sr
import pyttsx3
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from geopy.geocoders import Nominatim
import cv2
import paramiko
import warnings

# Suppressing warnings from paramiko
warnings.filterwarnings(action='ignore', module='.*paramiko.*')

# Initialize the recognizer and text-to-speech engine
recognizer = sr.Recognizer()
tts_engine = pyttsx3.init()
geolocator = Nominatim(user_agent="PythonSpeechRecognition")

def speak(text):
    """Function to provide voice feedback"""
    tts_engine.say(text)
    tts_engine.runAndWait()

def open_notepad():
    """Function to open Notepad"""
    try:
        subprocess.Popen(['notepad.exe'])
        speak("Opening Notepad")
    except FileNotFoundError:
        print("Notepad application not found.")
        speak("Notepad application not found.")

def open_chrome():
    """Function to open Chrome"""
    try:
        subprocess.Popen(['chrome.exe'])
        speak("Opening Chrome")
    except FileNotFoundError:
        print("Chrome application not found.")
        speak("Chrome application not found.")

def open_vm():
    """Function to open Oracle Virtual Machine"""
    try:
        subprocess.Popen([r'C:\\Program Files\\Oracle\\VirtualBox.exe'])  # Replace with the actual path to VirtualBox executable
        speak("Opening Oracle VM VirtualBox Manager")
    except FileNotFoundError:
        print("Oracle VM VirtualBox Manager application not found.")
        speak("Oracle VM VirtualBox Manager application not found.")

def show_location():
    """Function to show current location"""
    try:
        location = geolocator.geocode("address")
        print(f"Latitude: {location.latitude}, Longitude: {location.longitude}")
        speak(f"Latitude: {location.latitude}, Longitude: {location.longitude}")
    except Exception as e:
        print("Error retrieving location:", e)
        speak("Error retrieving location")

def sum_numbers(*args):
    """Function to sum numbers"""
    try:
        numbers = map(float, args)
        result = sum(numbers)
        print(f"The sum is: {result}")
        speak(f"The sum is {result}")
    except ValueError:
        print("Please provide valid numbers.")
        speak("Please provide valid numbers.")

def change_volume(direction):
    """Function to change the volume"""
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    current_volume = volume.GetMasterVolumeLevelScalar()
    if direction == "up":
        new_volume = min(current_volume + 0.1, 1.0)
    elif direction == "down":
        new_volume = max(current_volume - 0.1, 0.0)
    else:
        print("Invalid volume direction.")
        speak("Invalid volume direction.")
        return

    volume.SetMasterVolumeLevelScalar(new_volume, None)
    print(f"Volume {'increased' if direction == 'up' else 'decreased'}")
    speak(f"Volume {'increased' if direction == 'up' else 'decreased'}")

def start_video():
    """Function to start video capture"""
    cap = cv2.VideoCapture(0)
    while True:
        ret, frame = cap.read()
        cv2.imshow('Video', frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    cap.release()
    cv2.destroyAllWindows()

def capture_photo():
    """Function to capture a photo"""
    cap = cv2.VideoCapture(0)
    ret, frame = cap.read()
    cap.release()
    cv2.imwrite('captured_photo.jpg', frame)
    print("Photo captured successfully.")
    # Show the captured photo
    img = cv2.imread('captured_photo.jpg')
    cv2.imshow('Captured Photo', img)
    cv2.waitKey(0)
    cv2.destroyAllWindows()


def execute_ssh_command(host, username, password, command):
    """Function to execute SSH command"""
    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(host, username=username, password=password)
        stdin, stdout, stderr = ssh.exec_command(command)
        output = stdout.read().decode("utf-8")
        print(output)
        speak("SSH command executed successfully")
        ssh.close()
    except Exception as e:
        print("Error executing SSH command:", e)
        speak("Error executing SSH command")

def show_menu():
    """Function to display the menu"""
    print("\nAvailable commands:")
    print(" - open notepad")
    print(" - open chrome")
    print(" - open vm")
    print(" - show location")
    print(" - start video")
    print(" - capture photo")
    print(" - add <num1> <num2> ... <numN>")
    print(" - increase volume")
    print(" - decrease volume")
    print(" - execute SSH command")
    print(" - exit")

def parse_command(command):
    """Function to parse and execute the command"""
    tokens = shlex.split(command)
    if not tokens:
        return
   
    cmd = tokens[0].lower()
    args = tokens[1:]

    if cmd == "open" and args:
        if args[0].lower() == "notepad":
            open_notepad()
        elif args[0].lower() == "chrome":
            open_chrome()
        elif args[0].lower() == "vm":
            open_vm()
    elif cmd == "add":
        sum_numbers(*args)
    elif cmd == "show" and args[0].lower() == "location":
        show_location()
    elif cmd == "start" and args[0].lower() == "video":
        start_video()
    elif cmd == "capture" and args[0].lower() == "photo":
        capture_photo()
    elif cmd == "increase" and args[0].lower() == "volume":
        change_volume("up")
    elif cmd == "decrease" and args[0].lower() == "volume":
        change_volume("down")
    elif cmd == "execute":
        if len(args) == 4:
            host, username, password, ssh_command = args
            execute_ssh_command(host, username, password, ssh_command)
        else:
            print("Invalid number of arguments for SSH command.")
            speak("Invalid number of arguments for SSH command.")
    elif cmd == "exit":
        print("Thank you, Exiting")
        speak("Thank you, Exiting")
        return False
    else:
        print("Unknown command. Please try again.")
        speak("Unknown command. Please try again.")
   
    return True

def get_voice_command():
    """Function to capture voice input and return the command as text"""
    with sr.Microphone() as source:
        print("Listening for command...")
        speak("Listening for command")
        audio = recognizer.listen(source)
        try:
            command = recognizer.recognize_google(audio)
            print(f"You said: {command}")
            return command.lower()
        except sr.UnknownValueError:
            print("Sorry, I did not understand that.")
            speak("Sorry, I did not understand that.")
        except sr.RequestError:
            print("Sorry, there was an error with the speech recognition service.")
            speak("Sorry, there was an error with the speech recognition service.")
        return ""

def main():
    """Main function to run the menu"""
    print("Welcome to the Python Command Menu!")
    speak("Welcome to the Python Command Menu!")
    show_menu()
   
    while True:
        command = get_voice_command()
        if not parse_command(command):
            break

if __name__ == "__main__":
    main()

        
#FULL AND FINAL RUNNING CODE
