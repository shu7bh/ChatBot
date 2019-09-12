import pandas as pd
import string

readatabase = pd.read_excel('school.xlsx', sheet_name = 'Student Database')

Adno = readatabase['Adm No'].tolist()
Adno = Adno[len(Adno)- 1] - 1
Name = readatabase['Name'].tolist()
Class = readatabase['Class'].tolist()
Section = readatabase['Section'].tolist()
House = readatabase['House'].tolist()
Mname = readatabase['Mother Name'].tolist()
Fname = readatabase['Father Name'].tolist()
Contact = readatabase['Contact No'].tolist()
Mainsub = readatabase['Main Sub'].tolist()
Faddsub = readatabase['1st Add Sub'].tolist()
Saddsub = readatabase['2nd Add Sub'].tolist()
Transport = readatabase['Transport'].tolist()
Fees = readatabase['Fees'].tolist()
Message = readatabase['Messages'].tolist()
Assignment = readatabase['Assignments Pending'].tolist()
Progressreport = readatabase['Progress Report'].tolist()

def showdetails(num):
    if num > 0 and num < Adno + 2:
        num -= 1
        print("Bot: \tStudent Details\n\n" )
        print("\tName: " + Name[num])
        print("\tClass: ", Class[num])
        print("\tSection: " + Section[num])
        print("\tHouse: " + House[num])
        print("\tMother Name " + Mname[num])
        print("\tFather Name: " + Fname[num])
        print("\tContact Number: ", Contact[num], end = "\n\n")
        print("\tMain Subjects: ", end = ' ')
        print(*Mainsub[num].split(' '), sep = ', ')
        print("\tFirst Additional Subject: " + Faddsub[num])
        print("\tSecond Additional Subject: ", Saddsub[num], end = "\n\n")
        print("\tTransport availed: " + Transport[num])
        print("\tFees: ", Fees[num])
        print("\tNew Mesages: " + Message[num])
        print("\tNew Assignments: " + Assignment[num], end = "\n\n")
        print("\tProgress Report for first term: ")
        subjects = Mainsub[num].split(' ')
        subjects.append(Faddsub[num])
        subjects.append(Saddsub[num])
        grades = Progressreport[num].split(' ')
        for i in range(len(grades)):
            print("\t\t", subjects[i], ' - ', grades[i])
        return True
    else:
        print ("Bot: Wrong admission Number. Please enter again")
        return False
