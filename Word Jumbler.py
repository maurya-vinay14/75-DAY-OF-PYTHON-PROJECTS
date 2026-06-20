import random
import winsound
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

words = [
    "python",
    "computer",
    "programming",
    "developer",
    "keyboard",
    "monitor",
    "journey",
    "challenge",
    "project",
    "algorithm"
]

while True:
    score = 0
    rounds = 5

    print("WORD JUMBLER GAME 🔀")
    print(f"You have {rounds} rounds.")

    for round_num in range(1, rounds + 1):

        word = random.choice(words)

        jumbled = "".join(random.sample(word, len(word)))

        print(f"Round {round_num}")
        print("Jumbled Word:", jumbled)

        guess = input("Guess the original word: ").lower()

        if guess == word:
            print("Correct!")

            winsound.Beep(1000, 300)

            speaker.Speak("Correct Answer")

            score += 1

        else:
            print(f"Wrong! The word was: {word}")

            speaker.Speak("Wrong Answer")

    print("Game Over!")
    print(f"Your Score: {score}/{rounds}")

    if score == rounds:
        speaker.Speak("Excellent Performance")
    elif score >= 3:
        speaker.Speak("Good Job")
    else:
        speaker.Speak("Keep Practicing")

    play_again = input("Play Again? (y/n): ").lower()

    if play_again != "y":
        speaker.Speak("Thank you for playing")
        print("Thanks for playing!")
        break




























    