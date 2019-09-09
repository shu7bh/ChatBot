import sys
import pandas as pd
import string

readfaq = pd.read_excel('school.xlsx', sheet_name = 'FAQ')

question = readfaq['Question'].tolist()
response = readfaq['Response'].tolist()

readconversation = pd.read_excel('school.xlsx', sheet_name = 'Conversation')

greeting = readconversation['Greeting'].tolist()
praising = readconversation['Praising'].tolist()
ending = readconversation['Ending'].tolist()
cquestion = readconversation['Question'].tolist()
canswer = readconversation['Answer'].tolist()

res = "Hi"

def  Chatbot():
    while True:
        uinput = input("You: ").lower().translate(str.maketrans('', '', string.punctuation))

        if  any(phrase in uinput for phrase in ending):
            return

        elif any(phrase in uinput for phrase in praising):
            res = "Thanks a lot, always happy to help"

        elif any(phrase in uinput for phrase in greeting):
            if res == "Hi":
                res = "How may I help you?"
            else:
                res = "Hi, how may I assist you?"

        elif any(phrase in uinput for phrase in cquestion):
            for i in range (len(cquestion)):
                if cquestion[i] in uinput:
                    res = canswer[i]
                    break
        else:
            for i in range(len(question)):
                if question[i] in uinput:
                    res = response[i]
                    break
                elif i == len(question) - 1:
                    res = "Sorry couldn't able to get that, contact the reception for info regarding that"
                    file = open("questions.txt", 'a')
                    file.write(uinput + '\n')
                    file.close()
                    
        print ("Bot: " + res, end = "\n\n")

print ("Bot: Hi my name is Mark and I'm here to assist you?", end = "\n\n")
Chatbot()
print ("Bot: Bye, talk to you later. Hope I could help you out")
    
