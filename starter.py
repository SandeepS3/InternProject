# Sandeep Singh

# Libraries Needed
# import pandas as pd
import re
import openpyxl
import time
import random

# def gui:
# from tkinter import *
# canvas = Tk()

# canvas.geometry("500x250")

# label = Label(canvas, text = "How many roles are we assigning?")
# label.pack()

# numroles = Entry(canvas, width = 50)
# numroles.pack()

# button1 = Button(canvas)

# canvas.mainloop()


# Start

# Currently 19 roles total
lst = [
    "RSRTD",
    "RSRPS",
    "RSRWS",
    "RSROP",
    "Decanters",
    "DecantPS",
    "DecantWS",
    "DockRun",
    "DockCRun",
    "DockLL",
    "DockCPB",
    "DockDS",
    "DockPD",
    "DockLI",
    "DockIDRT",
    "DockUPPC",
    "DockPS",
    "DockCrets",
    "Clerk",
]
priolst = [
    "Clerk",
    "RSRTD",
    "RSROP",
    "RSRPS",
    "DecantPS",
    "DockPD",
    "DockIDRT",
    "DockPS",
    "DoctCrets",
]
usedppl = []


# Might still have to ask which day of the week it is
def startup():
    numofroles = input("How many roles will be used today? ")
    roles = {}

    # Running Through which roles are needed and how many for them
    for num in range(int(numofroles)):
        rolename = input("What is the role name? ")
        while True:
            if rolename in lst:
                break
            print("Invalid input, please try again!")
            time.sleep(1)
            rolename = input("What is the role name? ")
        pplrole = input("How many are needed for " + rolename + " today? ")
        roles[rolename] = pplrole
    return roles


# Reading Excel

# Reading the excel to create a dictionary


def retrieveperson(roles):
    def personlst():
        aalst = []

        # Going through the priority
        for row in range(2, spdsheet.max_row + 1):
            rls = spdsheet.cell(row, 4)
            rls1 = re.sub(",", "", str(rls.value))
            # Checks for all their prios and makes sure person not used already
            if (priolstloop in rls1.split()) and (row not in usedppl):
                aalst.append(row)
        return aalst

    # First loop through all the priority roles
    for priolstloop in priolst:
        # Check if this role is needed for the day, if it is continue on
        if priolstloop in roles:
            ss = openpyxl.load_workbook("test.xlsx")
            spdsheet = ss.active
            ppl = personlst()
            while roles[priolstloop] != 0:
                if len(ppl) == 0:
                    print(
                        "Not enough people for: "
                        + priolstloop
                        + ", still need "
                        + str(roles[priolstloop])
                        + " more people!"
                    )
                    break

                person = random.choice(ppl)

                # Second check to make sure they not already used
                if person not in usedppl:
                    spdsheet.cell(person, 8).value = priolstloop
                    usedppl.append(person)
                    roles[priolstloop] = int(roles[priolstloop]) - 1
                    ppl.remove(person)
                else:
                    ppl.remove(person)

            print(roles)
            ss.save("test.xlsx")


# Will now run for all the prio roles
# Need to ask is this before lunch or after so I can clear excel or not


r1 = startup()
retrieveperson(r1)
