import time
from datetime import datetime
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

print(" DIGITAL CLOCK ⌛")
speaker.Speak("Digital Clock Started")

while True:
    current_time = datetime.now().strftime("%H:%M:%S")

    print("\r Current Time:", current_time, end="")

    time.sleep(1)



