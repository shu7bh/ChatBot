import sys
import pandas as pd
import string
import StudentDetails as SD

readfaq = pd.read_excel('school.xlsx', sheet_name = 'FAQ')

question = readfaq['Question'].tolist()
response = readfaq['Response'].tolist()

readconversation = pd.read_excel('school.xlsx', sheet_name = 'Conversation')

greeting = readconversation['Greeting'].tolist()
praising = readconversation['Praising'].tolist()
ending = readconversation['Ending'].tolist()
cquestion = readconversation['Question'].tolist()
canswer = readconversation['Answer'].tolist()

uinput = " "
res = "Hi,"

def inCquestion():
    global uinput
    for i in range(len(cquestion)):
        if str(cquestion[i]) in uinput:
            print ("Bot: " + canswer[i])
            return True
    return False

def inFAQ():
    global uinput
    global res
    for i in range(len(question)):
        if question[i] in uinput:
            res = response[i]
            print ("Bot:", end = ' ')
            print(*res.split(' . '), sep = '\n        ')
            return True
    return False

def inGreeting():
    global uinput
    global res
           
    for phrase in greeting:
        if str(phrase) in uinput:
            res = str(phrase).capitalize() + ", how may I assist you?"
            print ("Bot: "+ res)
            return True
    return False

def inPraising():
    global uinput
    global res
    for phrase in praising:
        if str(phrase) in uinput:
            res = "Thanks a lot, always happy to help"
            print ("Bot: "+ res)
            return True
    return False
                
def  Chatbot():
    global res
    global uinput
    while True:
        print()
        uinput = input("You: ").lower().translate(str.maketrans('', '', string.punctuation))

        if res == "Please enter your admission number":
            if uinput.isdigit():
                if(SD.showdetails(int(uinput))):
                    res = ' '
                continue
            else:
                print("Bot: Wrong admission number. Please enter again.")
                continue
            
        if  any(phrase in uinput for phrase in ending):
            return
            
        if  inCquestion():
            continue
        if inFAQ():
            continue
        if inGreeting():
            continue
        if inPraising():
            continue

        file = open("questions.txt", 'a')
        file.write(uinput + '\n')
        file.close()
        print( "Bot: Sorry couldn't able to get that, contact the reception for info regarding that")

print("\n\n\t\t\t\t\t\tMark is the DPS-Modern Indian School's chatbot\n\n")
print ("\nBot: Hi, my name is Mark and I'm here to assist you?")
Chatbot()
print ("Bot: Bye, talk to you later. Hope I could help you out")
