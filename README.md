# Python-programs-with-Voice-Control

This Python script uses speech recognition and text-to-speech to perform various tasks on your computer, such as opening applications, changing volume, showing location, and executing SSH commands.

# Features
+ Open Applications: Open Notepad, Chrome, and Oracle VM VirtualBox Manager.
+ Show Location: Display the current location's latitude and longitude.
+ Change Volume: Increase or decrease the system volume.
+ Start Video: Start video capture using the webcam.
+ Capture Photo: Capture a photo using the webcam.
+ Execute SSH Command: Execute a command on a remote machine via SSH.
+ Voice-Controlled: Use voice commands to control the script.

# Requirements
+ Python 3.x
+ OpenCV (cv2)
+ pyttsx3
+ speech_recognition
+ geopy
+ paramiko
+ pycaw

# Installation
+ Clone the repository:
```git clone <repository_url>```

+ Install dependencies:
```pip install opencv-python pyttsx3 speechrecognition geopy paramiko pycaw ```

# Usage
Run the script:
```python voice_control.py```

# Available Commands
+ open notepad: Opens Notepad.
+ open chrome: Opens Chrome browser.
+ open vm: Opens Oracle VM VirtualBox Manager.
+ show location: Displays the current location's latitude and longitude.
+ start video: Starts video capture from the webcam.
+ capture photo: Captures a photo using the webcam.
+ add <num1> <num2> ... <numN>: Sums the provided numbers.
+ increase volume: Increases the system volume.
+ decrease volume: Decreases the system volume.
+ execute SSH command <host> <username> <password> <command>: Executes the provided SSH command on the specified host.
+ exit: Exits the application.

# Code Explanation
+ Import Libraries: Import necessary libraries including cv2, os, shlex, subprocess, speech_recognition, pyttsx3, geopy, paramiko, and pycaw.

+ Initialize Recognizer and TTS Engine: Initialize the speech recognizer and text-to-speech engine.

+ Voice Feedback Function: speak(text) provides voice feedback using pyttsx3.

+ Open Applications Functions: Functions to open Notepad (open_notepad()), Chrome (open_chrome()), and Oracle VM VirtualBox Manager (open_vm()).

+ Show Location Function: show_location() fetches and speaks the current location's latitude and longitude using geopy.

+ Sum Numbers Function: sum_numbers(*args) calculates and speaks the sum of the provided numbers.

+ Change Volume Function: change_volume(direction) adjusts the system volume using pycaw.

+ Start Video Function: start_video() captures video from the webcam and displays it.

+ Capture Photo Function: capture_photo() captures a photo using the webcam and displays it.

+ Execute SSH Command Function: execute_ssh_command(host, username, password, command) executes a command on a remote machine via SSH using paramiko.

+ Show Menu Function: show_menu() displays the available commands in the console.

+ Parse Command Function: parse_command(command) parses and executes the provided voice command.

+ Get Voice Command Function: get_voice_command() listens for a voice command and returns it as text.

+ Main Function: main() initializes the menu and processes voice commands in a loop until the exit command is given.
