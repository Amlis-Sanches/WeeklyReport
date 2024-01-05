'''
Description: This program is going to be the core application to pull information from the 
calander section and store it in a database. This data base will be used to seperate the
informaiton and generate a report for the user to send in an email. 
Start Date: 1/4/2024
Update Date: 1/4/2024

Notes:
To interact with Outlook and pull data from the calendar using Python, you would need the following tools:

1. `pywin32`: This is a Python library that provides some advanced utilities to interact 
with the Windows operating system. It includes modules for working with the Windows 
Registry, tasks, and other system features. In your case, you'll use it to interact 
with Microsoft Outlook through the COM interface.

2. `Microsoft Outlook`: You need to have Microsoft Outlook installed on your machine 
as `pywin32` will interact with it directly.

Here's a basic example of how you can use `pywin32` to interact with Outlook:

```python
import win32com.client

# Create an instance of the Outlook application
Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the calendar folder
calendar = Outlook.GetDefaultFolder(9).Items

# Loop through each item in the calendar
for appointment in calendar:
    print(appointment.Subject, appointment.Start, appointment.Duration)
```

This code will print out the subject, start time, and duration of each appointment in your default calendar.

Remember to install the `pywin32` library using pip:

```bash
pip install pywin32
```

Please note that this code will only work on a Windows machine with Outlook installed.
'''
import os
import sys
import time
import datetime as dt
import pandas as pd
import win32com.client
import tkinter as tk

def main():
    #user input
    check = False
    while check == False:
        startdate, check1 = datecheck("Enter the start date in the following format MM/DD/YYYY: ")
        enddate, check2 = datecheck("Enter the end date in the following format MM/DD/YYYY: ")
        if check1 == True and check2 == True:
            check = True
        else:
            print("Please enter the dates in the correct format")

    # Create an instance of the Outlook application
    Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Access the calendar folder
    calendar = Outlook.GetDefaultFolder(9).Items

    # Create a tkinter window
    window = tk.Tk()
    listbox = tk.Listbox(window)

    # Loop through each item in the calendar
    for appointment in calendar:
        listbox.insert(tk.END, f"{appointment.Subject}, {appointment.Start}, {appointment.Duration}, {appointment.Categories}")

    # Make the listbox expand and fill the entire window
    listbox.pack(expand=True, fill='both')

    window.mainloop()

    # Create a dataframe to store the calendar information
    df = pd.DataFrame(columns = ['Subject', 'Start', 'Duration', 'Categories'])

    # Loop through each item in the calendar
    for appointment in calendar:
        # Append the appointment data to the DataFrame
        df = df.append({
            'Subject': appointment.Subject,
            'Start': appointment.Start,
            'Duration': appointment.Duration,
            'Categories': appointment.Categories
        }, ignore_index=True)

    # Split the 'Start' column into two new columns 'Date' and 'Time'
    df['Date'], df['Time'] = df['Start'].str.split(' ', 1).str

    # Convert 'Date' to datetime format
    df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y')

    # If you want to convert 'Time' to a time format, you can do so like this:
    df['Time'] = pd.to_datetime(df['Time']).dt.time

def datecheck(prompt):
    while True:
        try:
            date = input(prompt)
            date = dt.datetime.strptime(date, '%m/%d/%Y')
            return date, True
        except ValueError:
            return None, False

if __name__ == "__main__":
    main()
