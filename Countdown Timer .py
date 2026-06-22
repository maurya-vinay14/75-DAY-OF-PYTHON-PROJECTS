import time
import winsound
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

while True:
    seconds = int(input("Enter countdown time in seconds: "))

    while seconds > 0:
        print(f"Time Remaining: {seconds} seconds")
        time.sleep(1)
        seconds -= 1

    print("Time's Up!")

    winsound.Beep(1000, 1000)

    speaker.Speak("Time is up")

    again = input("Start another timer? (y/n): ").lower()

    if again != "y":
        speaker.Speak("Goodbye")
        print("Goodbye!")
        break







    