import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

FILE_NAME = "notes.txt"

while True:
    print("NOTES APP")
    print("1. Add Note")
    print("2. View Notes")
    print("3. Delete All Notes")
    print("4. Exit")

    choice = input("Enter your choice: ")

    if choice == "1":
        note = input("Enter your note: ")

        with open(FILE_NAME, "a") as file:
            file.write(note + "\n")

        print("Note Saved Successfully!")
        speaker.Speak("Note Saved Successfully")

    elif choice == "2":
        try:
            with open(FILE_NAME, "r") as file:
                notes = file.readlines()

            if notes:
                print("Your Notes:")
                for i, note in enumerate(notes, start=1):
                    print(f"{i}. {note.strip()}")
            else:
                print("No Notes Found!")

        except FileNotFoundError:
            print("No Notes Found!")

    elif choice == "3":
        with open(FILE_NAME, "w") as file:
            pass

        print("All Notes Deleted!")
        speaker.Speak("All Notes Deleted")

    elif choice == "4":
        print("Exiting Notes App...")
        speaker.Speak("Goodbye")
        break

    else:
        print("Invalid Choice!")