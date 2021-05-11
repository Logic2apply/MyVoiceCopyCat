from win32com.client import Dispatch
import speech_recognition as sr
from pygame import mixer
import time
import os
import wave
import contextlib

def SayThis(pharase):
    speech = Dispatch("SAPI.SpVoice")
    speech.Speak(pharase)

while True:
    aud_recog = sr.Recognizer()
    with sr.Microphone() as source:
        # aud_recog.adjust_for_ambient_noise(source)  
        welcome_text = "Please Say Something..."
        print("******" + welcome_text + "******")
        SayThis(welcome_text)
        audio = aud_recog.listen(source)

    try:
        # recognize speech using Sphinx
        outputText = aud_recog.recognize_google(audio)
        print(f"\n\tI thinks you said {outputText}\n")

        # write audio to a WAV file
        aud_file = "myvoicecopycat.wav"
        with open(aud_file, "wb") as f:
            f.write(audio.get_wav_data())
        
        mixer.init()
        mixer.music.load(aud_file)
        mixer.music.play()
        
        with contextlib.closing(wave.open(aud_file,'r')) as f:
            frames = f.getnframes()
            rate = f.getframerate()
            duration = frames / float(rate)
        time.sleep(duration)
        mixer.music.pause()
        mixer.music.unload()
        os.remove(aud_file)
        
        # Attempt 1
        # outputText = aud_recog.recognize_google(audio, language="hi-IN")
        # print(f"I thinks you said {outputText}")
        # SayThis(outputText)

    except sr.UnknownValueError:
        print("I could not understand your voice... Can You say this again?")
        SayThis("I could not understand your voice... Can You say this again?")

    except sr.RequestError as e:
        print(f"I error; {e}")
        SayThis(f"I error; {e}")