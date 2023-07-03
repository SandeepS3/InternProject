# Sandeep Singh

# Libraries Needed
# import pandas as pd
import openpyxl
# import re
import time

# namelst = []
# #dataframe1 = pd.read_excel('test.xlsx')
# dataframe = openpyxl.load_workbook("test.xlsx")

# dataframe1 = dataframe.active
# #names into list
# for name in range(2, dataframe1.max_row+1):
#     names = dataframe1.cell(name,2)
#     namelst.append(names.value)

# for j in range(2, dataframe1.max_row+1):
#     answer = dataframe1.cell(j,4)
#     nanswer = re.sub(",", "", str(answer.value))
#     if "Stow" in nanswer.split():
#         print(namelst[j-2])
#         print(nanswer.split())
# print(namelst)

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

# Currently 19 roles
lst = ["RSRTD", "RSRPS", "RSRWS", "RSROP", "Decanters", "DecantPS", "DecantWS", "DockRun", "DockCRun",
       "DockLL", "DockCPB", "DockDS", "DockPD", "DockLI", "DockIDRT", "DockUPPC", "DockPS", "DockCrets", "Clerk"]
priolst = ["Clerk", "RSRTD", "RSROP", "DecantPS", "DockPD", "DockIDRT", "DockPS", "DoctCrets"]
# Might still have to ask which day of the week it is


def startup():
    numofroles = input("How many roles will be used today? ")
    roles = {}

    # Running Through which roles are needed and how many for them
    for num in range(int(numofroles)):
        rolename = input("What is the role name? ")
        while (True):
            if (rolename in lst):
                break
            print("Invalid input, please try again!")
            time.sleep(1)
            rolename = input("What is the role name? ")
        pplrole = input("How many are needed for " + rolename + " today? ")
        roles[rolename] = pplrole
    # print(roles)

# Reading Excel

# Reading the excel to create a dictionary


def reading():
    spdsheet = (openpyxl.load_workbook("test.xlsx")).active


startup()
