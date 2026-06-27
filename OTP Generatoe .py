import random
import time
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

print("OTP GENERATOR 🔐")

length = int(input("Enter OTP Length (4-8): "))

if length < 4 or length > 8:
    print("OTP length should be between 4 and 8 digits.")
    speaker.Speak("Invalid OTP Length")

else:
    # Generate OTP
    otp = str(random.randint(1, 9))

    for i in range(length - 1):
        otp += str(random.randint(0, 9))

    print(f"Your OTP is: {otp}")
    speaker.Speak("OTP Generated Successfully")

    # OTP expires after 30 seconds
    expiry_time = time.time() + 30

    attempts = 3

    while attempts > 0:

        remaining_time = int(expiry_time - time.time())

        if remaining_time <= 0:
            print("OTP Expired!")
            speaker.Speak("OTP Expired")
            break

        print(f"Time Remaining: {remaining_time} seconds")

        entered = input("Enter OTP: ")

        if entered == otp:
            print("OTP Verified Successfully!")
            speaker.Speak("OTP Verified Successfully")
            break

        else:
            attempts -= 1
            print(f"Incorrect OTP! Attempts Left: {attempts}")

    if attempts == 0:
        print("Verification Failed!")
        speaker.Speak("Verification Failed")