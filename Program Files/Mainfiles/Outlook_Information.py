'''
This file is for creating the main outlook code to extract the data when needed from the outlook file. I will have two parts of the code:
1. pulling all outlook data from the calendar - Returns a dataframe
2. pull a selected amount of data - returns a calandar
'''


import datetime as dt
import pandas as pd
import win32com.client
import tkinter as tk

def main():
    pass

def pull_all_OL_data():
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
    df = transform_to_df(listbox)
    return df

def pull_select_OL_data(start_date, end_date):
    # Create an instance of the Outlook application
    Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Access the calendar folder
    calendar = Outlook.GetDefaultFolder(9).Items

    # Set the start and end dates for the date range
    start_date = dt.date(2022, 1, 1)
    end_date = dt.date(2022, 12, 31)

    # Format the dates in the format expected by Outlook
    start_date = start_date.strftime("%m/%d/%Y")
    end_date = end_date.strftime("%m/%d/%Y")

    # Create the restriction string
    restriction = "[Start] >= '{0}' AND [Start] <= '{1}'".format(start_date, end_date)

    # Apply the restriction (filter the items)
    calendar = calendar.Restrict(restriction)

    # Create a tkinter window
    window = tk.Tk()
    listbox = tk.Listbox(window)

    # Loop through each item in the calendar
    for appointment in calendar:
        listbox.insert(tk.END, f"{appointment.Subject}, {appointment.Start}, {appointment.Duration}, {appointment.Categories}")

    # Make the listbox expand and fill the entire window
    listbox.pack(expand=True, fill='both')

    window.mainloop()
    df = transform_to_df(listbox)
    return df

def transform_to_df(OL_data):
    # Loop through each item in the calendar
    for appointment in OL_data:
        # Append the appointment data to the DataFrame
        df = df.append({
            'Subject': appointment.Subject,
            'Start': appointment.Start,
            'Duration': appointment.Duration,
            'Categories': appointment.Categories
        }, ignore_index=True)
    return df

if __name__ == "__main__":
    main()