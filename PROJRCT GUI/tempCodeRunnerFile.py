import tkinter as tk
import pandas as pd
import os

def create_database():
    area = entry.get()

    # Connect to database and create database for the specified area

    message_label.config(text="Database for area '" + area + "' created successfully!")

    open_button = tk.Button(root, text="Open Database", command=open_database)
    open_button.pack()

def open_database():
    # Open the database for the specified area

    message_label.config(text="Database for area opened successfully!")

    create_table_button = tk.Button(root, text="Create Table from Excel", command=create_table)
    create_table_button.pack()

def create_table():
    # Open a new window to select the Excel file

    file_path = "C:\\Users\\ISHANK\\Desktop\\PROJRCT GUI\\DATABASEAREA.xlsx"

    if file_path:
        # Read the Excel file into a pandas DataFrame

        data = pd.read_excel(file_path)

        # Open the specified Excel file

        os.startfile(file_path)

def next_step():
    # Implement the next step of the process, such as creating the table in the database

    message_label.config(text="Table created successfully!")

root = tk.Tk()
root.title("Create Database")

label = tk.Label(root, text="Enter the area in which the building is located:")
label.pack()

entry = tk.Entry(root)
entry.pack()

submit_button = tk.Button(root, text="Submit", command=create_database)
submit_button.pack()

message_label = tk.Label(root, text="Click 'Submit' to create a database for the above area.")
message_label.pack()

root.mainloop()