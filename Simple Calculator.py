import winsound
import win32com.client
print("🧮 Simple Calculator")

num1 = float(input("Enter first number: "))
operator = input("Enter operator (+, -, *, /): ")
num2 = float(input("Enter second number: "))


winsound.Beep(800,250)

if operator == "+":
    print("Result =", num1 + num2)

elif operator == "-":
    print("Result =", num1 - num2)

elif operator == "*":
    print("Result =", num1 * num2)

elif operator == "/":
    if num2 != 0:
        print("Result =", num1 / num2)
    else:
        print("Error: Cannot divide by zero!")

else:
    print("Invalid operator!")

    
    
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Speak("Solved")
