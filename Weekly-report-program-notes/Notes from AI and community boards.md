Developing your program in Python is definitely feasible, especially if you are more comfortable with Python than with VB.NET. Python has libraries that can interface with Outlook and Google Calendar, as well as libraries for data manipulation and creating Excel files.

Here are the steps and Python libraries you might consider using:

1. **Outlook Integration**: You can use the `pywin32` library to interact with Outlook via Python. This allows you to access, read, and manipulate data in Outlook.
    
2. **Excel File Creation and Manipulation**: Libraries like `openpyxl` or `pandas` are great for creating and manipulating Excel files in Python. `openpyxl` is more focused on Excel files specifically, while `pandas` is a more comprehensive data analysis library that can also handle Excel files.
    
3. **Google Calendar Integration**: For Google Calendar, you can use Google's official `google-api-python-client` library. It allows you to interact with Google Calendar and other Google services.
    
4. **Cross-platform Synchronization**: For the more advanced feature of synchronizing calendars across different platforms, you'll need to interface with the APIs of each platform (Outlook, Google Calendar, etc.) and implement logic to sync data among them.
    
5. **Android App Development**: If you plan to implement this into an Android app, you might need to switch to a language like Kotlin or Java, which are standard for Android app development. However, you can use Python scripts for backend logic.
    
6. **Ease of Use and Refinement**: Python's extensive libraries for GUI development (like `Tkinter`, `PyQt`, or `Kivy`) can help in creating user interfaces for ease of use. Additionally, Python's flexibility and readability can aid in refining and updating the program.
    
7. **Visual Studio as IDE**: You can use Visual Studio for Python development, as it supports Python through extensions. Visual Studio provides a robust environment for Python development, including debugging and project management.
    

Remember, while Python can interact with these services, the specifics of how to read from one calendar, interpret the data, and then format and send it to another system can get complex, especially when dealing with multiple different platforms. Each step will likely involve quite a bit of custom programming.

Lastly, always ensure you're handling any sensitive data (like calendar and email details) securely and in compliance with any relevant privacy laws or policies.