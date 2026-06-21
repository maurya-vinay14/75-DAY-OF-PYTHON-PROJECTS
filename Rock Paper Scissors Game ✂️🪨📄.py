import random
import winsound
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

choices = ["rock", "paper", "scissors"]

while True:
    player_score = 0
    computer_score = 0

    print("ROCK PAPER SCISSORS 🎮")
    print("First to 3 wins!\n")

    while player_score < 3 and computer_score < 3:

        print("1. Rock")
        print("2. Paper")
        print("3. Scissors")

        choice = input("Enter your choice (1/2/3): ")

        if choice == "1":
            player = "rock"
        elif choice == "2":
            player = "paper"
        elif choice == "3":
            player = "scissors"
        else:
            print("Invalid Choice!")
            continue

        computer = random.choice(choices)

        print(f"You: {player}")
        print(f"Computer: {computer}")

        if player == computer:
            print("It's a Draw!")
            speaker.Speak("It's a Draw")

        elif (
            (player == "rock" and computer == "scissors")
            or (player == "paper" and computer == "rock")
            or (player == "scissors" and computer == "paper")
        ):
            print("You Win This Round!")
            winsound.Beep(1000, 300)
            speaker.Speak("You Win")
            player_score += 1

        else:
            print("Computer Wins This Round!")
            speaker.Speak("Computer Wins")
            computer_score += 1

        print("Scoreboard")
        print(f"You: {player_score}")
        print(f"Computer: {computer_score}")
        print("-" * 25)

    if player_score == 3:
        print("Congratulations! You Won The Match!")
        speaker.Speak("Congratulations. You Won The Match")

    else:
        print("Computer Won The Match!")
        speaker.Speak("Computer Won The Match")

    play_again = input("Play Again? (y/n): ").lower()

    if play_again != "y":
        speaker.Speak("Thank You For Playing")
        print("Thanks For Playing!")
        break