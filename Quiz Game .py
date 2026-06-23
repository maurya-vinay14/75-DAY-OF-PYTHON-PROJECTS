import win32com.client
import winsound

speaker = win32com.client.Dispatch("SAPI.SpVoice")

questions = [
    {
        "question": "What does CPU stand for?",
        "answer": "central processing unit"
    },
    {
        "question": "Which language are we using in this project?",
        "answer": "python"
    },
    {
        "question": "What is 5 + 7?",
        "answer": "12"
    },
    {
        "question": "Who developed Python?",
        "answer": "guido van rossum"
    },
    {
        "question": "What keyword is used to create a loop that runs while a condition is true?",
        "answer": "while"
    }
]

score = 0

print("Welcome to the Quiz Game! ")

for q in questions:

    print(q["question"])

    user_answer = input("Your Answer: ").lower()

    if user_answer == q["answer"]:
        print("Correct! ")

        winsound.Beep(1000, 300)

        speaker.Speak("Correct Answer")

        score += 1

    else:
        print(f"Wrong! Correct answer: {q['answer']}")

        speaker.Speak("Wrong Answer")

print("Quiz Completed!")
print(f"Your Score: {score}/{len(questions)}")

if score == len(questions):
    speaker.Speak("Excellent Performance")
elif score >= 3:
    speaker.Speak("Good Job")
else:
    speaker.Speak("Keep Practicing")