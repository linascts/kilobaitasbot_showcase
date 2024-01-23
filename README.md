# KilobaitasBOT description

KilobaitasBOT is used to save time by automating repetetive tasks in a legacy software. In total this software saves 48 hours of work every month for an employee.

It was created mostly using Pywinauto, Pyautogui, Openpyxl, Selenium libraries for automation and Tkinter library for GUI. Ideally using pywinauto let us connect to the backend of the legacy software for a more solid automation, however a more primitve solution like Pyautogui image recognition was used for navigation due to failures connecting to the backend. 

There is three integrated modules so far that do different tasks, the user can freely change the options in these modules, for example if he would like to gather a new manufacturer's data or adding a new type of Graphics Card to filter in the software, the user could change the settings and the information would be saved to JSON.

KilobaitasBOT works like this:

1. The user changes the settings as he sees fit, and writes down the login information of the legacy software so that the program could copy-paste it if the legacy software itself is not running at the moment.

2. The software at first checks it's initial enviroment, how many monitors are connected, language settings, legacy software version, if there's any programs workking that interferes with the autoamtion process, and so on. The software adjusts itself for the enviroment or stops the sequence if there's some kind of error.

3. After the initial phase, the software will login into the legacy software and does the repetitive tasks depending on the module selected.

4. After extracting necessary data from the legacy software, the software continues to create reports with Excel so it could be sent to manufacturers or for the company own use.

![GUI-1](gif/KilobaitasBot-1.gif)

![GUI-2](gif/KilobaitasBot-2.gif)



