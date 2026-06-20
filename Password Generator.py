import random
import win32com.client

chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!@#$%^&*()"

# Voice engine
speaker = win32com.client.Dispatch("SAPI.SpVoice")


length = int(input("Enter length of password: "))
password  = ""

# Password strength
if length < 8:
    print("🔴 Password Strength: Weak")
elif length < 12:
    print("🟡 Password Strength: Medium")
else:
    print("🟢 Password Strength: Strong")

print("Generated Password(s):")

for a in range(length):
    password += random.choice(chars)
print(password)

# Voice notification
speaker.Speak("Password Generated Successfully " )
