# PHASE 1: 
# THIS PHASE CONTAINS BASIC VOICE COMMAND RECOGNITION

import argparse
import speech_recognition as sr  # Library for speech recognition
import os
import webbrowser
import subprocess
import pyttsx3
from fuzzywuzzy import fuzz
import psutil
from datetime import datetime
try:
    import noisereduce as nr
    import numpy as np
    _HAS_NOISEREDUCE = True
except Exception:
    # noisereduce can be slow to import or unavailable; make it optional
    _HAS_NOISEREDUCE = False
    np = None

import threading
import time

# Initialize recognizer and text-to-speech engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()
engine.setProperty('rate', 135)  # set speech speed

print("Phase 1: Basic speech recognition")


def _speak_worker(text):
    """Worker that runs TTS on a separate thread to avoid blocking."""
    try:
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
        # swallow TTS errors so they don't crash the assistant loop
        print("TTS error:", e)


def speak(text, block=True):
    """Text-to-Speech Function. When block=False, runs in background thread."""
    if block:
        _speak_worker(text)
    else:
        t = threading.Thread(target=_speak_worker, args=(text,), daemon=True)
        t.start()


def reduce_noise(audio_data):
    """Reduce noise in the captured audio using noisereduce library.

    Returns a numpy array of reduced audio samples. If noisereduce is not
    available, raises RuntimeError.
    """
    if not _HAS_NOISEREDUCE:
        raise RuntimeError("noisereduce not available")

    # Convert audio to numpy array
    np_audio = np.frombuffer(audio_data.get_raw_data(), dtype=np.int16)

    # Apply noise reduction
    reduced_audio = nr.reduce_noise(y=np_audio, sr=audio_data.sample_rate)
    return reduced_audio


# Function to open any app dynamically
def open_app(app_name):
    """
    Open any application dynamically by its name using os.system('start').
    """
    app_name = app_name.lower()  # Normalize app name for consistency

    if "notepad" in app_name:
        os.system("start notepad")
        return "Notepad opened."

    elif "calculator" in app_name:
        os.system("start calc")
        return "Calculator opened."

    elif "powerpoint" in app_name:
        os.system("start powerpnt")
        return "PowerPoint opened."

    elif "word" in app_name:
        os.system("start winword")
        return "Word opened."

    else:
        return f"Sorry, I couldn't find an application named {app_name}."


def close_app(app_name):
    """
    Close the application based on its process name.
    """
    app_name = app_name.lower()

    # Map user-friendly app names to process names
    process_map = {
        "notepad": "notepad.exe",
        "calculator": "Calculator.exe",
        "powerpoint": "POWERPNT.EXE",
        "word": "WINWORD.EXE"
    }

    process_name = process_map.get(app_name)  # Get the actual process name

    if process_name:
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'].lower() == process_name.lower():
                proc.kill()  # Kill the process
                return f"{app_name.capitalize()} closed successfully."
        return f"{app_name.capitalize()} is not running."
    else:
        return f"I don't know how to close {app_name}."


# Main command handling logic
def process_command(command):
    command = command.lower()  # Normalize the command for matching

    # Match 'open' commands
    if "open" in command:
        app_name = command.replace("open", "").strip()
        if app_name:
            response = open_app(app_name)
        else:
            response = "I didn't catch the application name. Please try again."
        print(response)
        speak(response)

    # Match 'close' or 'terminate' commands
    elif "close" in command or "terminate" in command:
        # Extract the app name dynamically
        app_name = command.replace("close", "").replace("terminate", "").strip()
        if app_name:
            response = close_app(app_name)
        else:
            response = "I didn't catch the application name. Please try again."
        print(response)
        speak(response)

    # Match time-related queries
    elif fuzz.partial_ratio(command, "what is the time") > 80 or fuzz.partial_ratio(command, "tell the time") > 80:
        current_time = datetime.now().strftime("%I:%M %p")
        response = f"The time is {current_time}"
        print(response)
        speak(response)

    # Match exit/quit commands
    elif fuzz.partial_ratio(command, "exit") > 80 or fuzz.partial_ratio(command, "quit") > 80:
        response = "Goodbye, see you again!"
        print(response)
        speak(response)
        exit()  # Exit the program

    # Search functionality
    elif "search" in command or "look up" in command:
        # Extract the query
        query = command.replace("search for", "").replace("look up", "").replace("search", "").strip()
        if query:
            response = f"Searching for {query} on the web."
            print(response)
            speak(response)
            webbrowser.open(f"https://www.google.com/search?q={query}")

    # Default fallback response
    else:
        response = "Command not recognized. Please try again."
        print(response)
        speak(response)


def main(args=None):
    """Main entry point for the voice assistant.

    Args:
        args: argparse.Namespace with attributes 'fast' and 'noisereduce' and 'startup_test'
    """
    parser = None
    if args is None:
        parser = argparse.ArgumentParser()
        parser.add_argument('--fast', action='store_true', help='Enable fast-mode optimizations')
        parser.add_argument('--noisereduce', action='store_true', help='Force enable noise reduction (may be slow)')
        parser.add_argument('--startup-test', action='store_true', help='Run only startup/init and exit (for timing)')
        args = parser.parse_args()

    # Calibration duration: shorter in fast mode
    calib_duration = 0.5 if getattr(args, 'fast', False) else 2.0
    use_noisereduce = getattr(args, 'noisereduce', False) and _HAS_NOISEREDUCE

    print(f"Starting assistant (fast={getattr(args, 'fast', False)}, noisereduce={use_noisereduce})")

    if getattr(args, 'startup_test', False):
        # Quick startup test - do initialization then exit
        t0 = time.time()
        # Perform a quick microphone open/close to ensure permissions
        try:
            with sr.Microphone() as source:
                recognizer.adjust_for_ambient_noise(source, duration=calib_duration)
        except Exception as e:
            print('Microphone init warning:', e)
        print('Startup init time:', time.time() - t0)
        return

    # ENABLING CONTINUOUS ITERATIVE LISTENING FOR MULTIPLE COMMAND EXECUTIONS
    while True:
        try:
            with sr.Microphone() as source:
                print("\nListening... (Say 'exit' to quit)")
                recognizer.adjust_for_ambient_noise(source, duration=calib_duration)  # Dynamically adjust for background noise
                # Use a phrase_time_limit in fast mode so the listen call returns quickly
                phrase_limit = 5 if getattr(args, 'fast', False) else None
                audio = recognizer.listen(source, phrase_time_limit=phrase_limit)  # Capture audio input

                # Optionally reduce noise using noisereduce (can be slow)
                audio_data = audio
                if use_noisereduce:
                    try:
                        reduced_audio = reduce_noise(audio)
                        audio_data = sr.AudioData(reduced_audio.tobytes(), audio.sample_rate, audio.sample_width)
                    except Exception as e:
                        print("Skipping noise reduction due to an error:", e)
                        audio_data = audio

                print("Recognizing...")

                # Recognize speech using Google API
                try:
                    text = recognizer.recognize_google(audio_data)
                    print(f"You said: {text}")
                    process_command(text)  # Process the recognized command
                except sr.UnknownValueError:
                    error_msg = "Sorry, I could not understand the audio."
                    print(error_msg)
                    speak(error_msg, block=False)
                except sr.RequestError as e:
                    error_msg = f"Could not request results; {e}"
                    print(error_msg)
                    speak(error_msg, block=False)

        except KeyboardInterrupt:
            goodbye_msg = "Program terminated manually."
            print(goodbye_msg)
            speak(goodbye_msg)
            break


if __name__ == '__main__':
    main()
