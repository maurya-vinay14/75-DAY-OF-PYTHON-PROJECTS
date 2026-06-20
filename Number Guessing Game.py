import random
import winsound
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

while True:
    print("Number Guessing Game 🎮")

    print("Select Difficulty:")
    print("1. Easy (1-10)")
    print("2. Medium (1-50)")
    print("3. Hard (1-100)")

    choice = input("Enter your choice (1/2/3): ")

    if choice == "1":
        limit = 10
    elif choice == "2":
        limit = 50
    elif choice == "3":
        limit = 100
    else:
        print("Invalid Choice! Defaulting to Hard Mode.")
        limit = 100

    secret_number = random.randint(1, limit)
    attempts = 0

    print(f"I have chosen a number between 1 and {limit}.")
    print("Try to guess it!")

    while True:
        guess = int(input("Enter your guess: "))
        attempts += 1

        if guess < secret_number:
            print("Too Low!📉")
        elif guess > secret_number:
            print("Too High!📈")
        else:
            print(f" Congratulations! You guessed the number in {attempts} attempts.🎉")

            winsound.Beep(1000, 500)

            speaker.Speak("Congratulations, you guessed the number")

            break

    play_again = input("Play Again? (y/n): ").lower()

    if play_again != "y":
        speaker.Speak("Thank you for playing")
        print("hanks for playing!")
        break