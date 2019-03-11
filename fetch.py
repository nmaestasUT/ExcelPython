import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import xlrd
from tkinter import *




nameOfFile = input("Enter File Name: ")
nameOfSheet = input("Enter Sheet Name: ")
currentMonth = input("Enter Month: ")


totalPrice = 0

file = nameOfFile
df = pd.read_excel(file, sheet_name=nameOfSheet)

#get name from user and create dataframe that matches from column named "name"
nameOfStudent = input("name: ")
df1 = df[df['Name'] == nameOfStudent]

#save name on variable x, concatenate strings.
x = "Student name: " + nameOfStudent + " \n Month: " + currentMonth

#for every row in fetched dataframe, save string in variable x
for row in df1.itertuples():
    if (row[6]!="credit"):
        #fetch specified columns for each row
        x += row[6] + " (" + row[9] + ") " + row[8] + " - " + str(abs(row[10])) + " => " + str(format(int(row[16]), ','))
    x += "\n"
    if (row[6]=="credit"):
        x += "Credit " + row[8] + " " + str(abs(row[10])) + " => " + str(format(int(row[16]), ','))
    totalPrice += row[16]

#
print(x, "\n");

totalPrice = format(int(totalPrice), ',')
