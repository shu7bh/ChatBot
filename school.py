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

res = "Hi,"

def inCquestion(q):
    for i in range(len(cquestion)):
        if str(cquestion[i]) in q:
            print ("Bot: " + canswer[i])
            return True
    return False

def inFAQ(q):
    global res
    for i in range(len(question)):
        if question[i] in q:
            res = response[i]
            print ("Bot: " + res)
            return True
    return False
                
def  Chatbot():
    global res
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
            
        if  inCquestion(uinput):
            continue

        if inFAQ(uinput):
            continue

        if "hi" in uinput:
            if "Hi," in res:
                res = "How may I help you?"
            else:
                res = "Hi, how may I assist you?"
            print ("Bot: "+ res)
            continue
        
        if any(phrase in uinput for phrase in greeting):          
            if "Hi," in res:
                res = "How may I help you?"
            else:
                res = "Hi, how may I assist you?"
            print ("Bot: "+ res)
            continue
            
        if any(phrase in uinput for phrase in praising):
            res = "Thanks a lot, always happy to help"
            print ("Bot: "+ res)
            continue

        file = open("questions.txt", 'a')
        file.write(uinput + '\n')
        file.close()
        print( "Bot: Sorry couldn't able to get that, contact the reception for info regarding that")
        #print ("Bot: " + res, end = "\n\n")

print ("\nBot: Hi, my name is Mark and I'm here to assist you?")
Chatbot()
print ("Bot: Bye, talk to you later. Hope I could help you out")
    
