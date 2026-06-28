import re
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

valid = 0
invalid = 0

print("EMAIL VALIDATOR 📧")

pattern = r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'

while True:

    email = input("\nEnter Email (or type 'exit'): ")

    if email.lower() == "exit":
        break

    if re.fullmatch(pattern, email):

        print("Valid Email")

        if "@gmail.com" in email:
            print("📨 Provider : Gmail")

        elif "@outlook.com" in email:
            print("📨 Provider : Outlook")

        elif "@yahoo.com" in email:
            print("📨 Provider : Yahoo")

        else:
            print("📨 Provider : Other")

        with open("emails.txt", "a") as file:
            file.write(email + "\n")

        speaker.Speak("Valid Email Address")

        valid += 1

    else:
        print("Invalid Email")

        speaker.Speak("Invalid Email Address")

        invalid += 1

print("SUMMARY")
print(f"Valid Emails   : {valid}")
print(f"Invalid Emails : {invalid}")

speaker.Speak("Program Closed")