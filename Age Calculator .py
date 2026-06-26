from datetime import datetime
import win32com.client

# Voice Engine
speaker = win32com.client.Dispatch("SAPI.SpVoice")

print("AGE CALCULATOR 🎂")

try:
    # User Input
    dob = input("Enter your Date of Birth (DD/MM/YYYY): ")

    # Convert String to Date
    birth_date = datetime.strptime(dob, "%d/%m/%Y")

    # Current Date
    today = datetime.today()

    # Calculate Age
    age = today.year - birth_date.year

    # Check if birthday has occurred this year
    if (today.month, today.day) < (birth_date.month, birth_date.day):
        age -= 1

    # Total Days
    total_days = (today - birth_date).days

    # Approximate Months
    total_months = total_days // 30

    # Hours
    total_hours = total_days * 24

    # Minutes
    total_minutes = total_hours * 60

    # Seconds
    total_seconds = total_minutes * 60

    # Next Birthday
    next_birthday = birth_date.replace(year=today.year)

    if next_birthday < today:
        next_birthday = next_birthday.replace(year=today.year + 1)

    days_left = (next_birthday - today).days

    # Output
    print(" AGE DETAILS 🎉")
    print("-" * 35)
    print(f" Age              : {age} Years")
    print(f" Days Lived       : {total_days}")
    print(f" Months Lived     : {total_months} (Approx.)")
    print(f" Hours Lived      : {total_hours}")
    print(f" Minutes Lived    : {total_minutes}")
    print(f" Seconds Lived    : {total_seconds}")
    print(f" Next Birthday In : {days_left} Days")
    print("-" * 35)

    # Voice Output
    speaker.Speak(f"You are {age} years old.")
    speaker.Speak(f"Your next birthday is in {days_left} days.")

except ValueError:
    print("Invalid Date Format!")
    print("Please enter your date in DD/MM/YYYY format.")
    speaker.Speak("Invalid Date Format. Please try again.")