import win32com.client

# Create an instance of the Outlook application
Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access the calendar folder (9 is the code for the calendar folder)
calendar = Outlook.GetDefaultFolder(9).Items

# Loop through each item in the calendar
for appointment in calendar:
    print(f"Subject: {appointment.Subject}, Start: {appointment.Start}, Duration: {appointment.Duration}")