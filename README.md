# KilobaitasBot showcase

To install .exe file, use PyInstaller in cmd:

pyinstaller --onedir -w --name KilobaitasBot --icon "images\gui_images\kilobaitas-bot-icon.ico" --add-data "C:\Users\linas\AppData\Local\Programs\Python\Python311-32\Lib\site-packages\customtkinter;customtkinter" UserControl.py


## Software description


This software is used to save time by automating repetetive tasks in a legacy software, at first the automation was working with Microsoft Power Automate, however it was quite hard to trace down errors and was very limited in scope compared to Python. In total this software saves 48 hours of work every month for an employee.

It was created mostly using Pywinauto, Pyautogui, Openpyxl, Selenium libraries for automation and Tkinter library for GUI. Ideally using pywinauto let us connect to the backend of the software for a more solid automation, however a more primitve solution like Pyautogui image recognition was used for navigation due to failures connecting to the backend.

There is three integrated modules so far that do different tasks, the user can freely change the options in these modules, for example if he would like to gather a new manufacturer's data or adding a new type of Graphics Card to filter in the software, the user could change the settings and the information would be saved to JSON.

![GUI-1](gif/KilobaitasBot-1.gif)

![GUI-2](gif/KilobaitasBot-2.gif)



