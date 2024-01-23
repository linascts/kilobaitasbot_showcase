import warnings
warnings.filterwarnings("ignore") # supress all warnings
import sys
sys.coinit_flags = 2 # pyinwauto / tkinter thread freeze solution
import WeeklyReportModule as wrp
import WeeklySpecReportModule as wmrp
import MonthlySpecReportModule as mmrp
import tkinter as tk
import customtkinter, os, json, threading
import datetime
from PIL import Image
from tkinter import filedialog
from pywinauto import application

# Global Variables
spec_component_validation = False
spec_component_old_validation_error = None
spec_component_clicked_before = False
old_anchor_frame = None
old_new_component_label = None
old_new_component_entry = None
old_new_component_button = None
old_error_label_2 = None
old_save_button_2 = None
old_back_button_2 = None

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.window_width = 700
        self.window_height = 550
        self.screen_width = self.winfo_screenwidth()
        self.screen_height = self.winfo_screenheight()
        self.center_x = int(self.screen_width/2 - self.window_width / 2)
        self.center_y = int(self.screen_height/2 - self.window_height / 2)
        self.title("KilobaitasBot")
        self.geometry(f'{self.window_width}x{self.window_height}+{self.center_x}+{self.center_y}')
        self.resizable(0, 0)
        self.iconbitmap("images/gui_images/kilobaitas-bot-icon.ico")

        f = open("data/DirectoryData.json", encoding='utf-8')
        f_data = json.load(f)
        all_directories = f_data.get("directories")
        report_directory = all_directories.get("report_directory")
        f.close()

        default_export_text = tk.StringVar()
        default_export_text.set(report_directory)
        
        # set grid layout 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # load images with light and dark mode image
        image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "images\gui_images")
        self.logo_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "kilobaitas_logo.png")), size=(150, 25))
        self.large_test_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "large_test_image.png")), size=(500, 150))
        self.image_icon_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "image_icon_light.png")), size=(20, 20))
        self.frame_1_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "7_dark.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "7_light.png")), size=(20, 20))
        self.frame_2_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "7_dark.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "7_light.png")), size=(20, 20))
        self.frame_3_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "31_dark.png")),
                                                     dark_image=Image.open(os.path.join(image_path, "31_light.png")), size=(20, 20))
        

        # create navigation frame
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(4, weight=1)

        self.navigation_frame_label = customtkinter.CTkLabel(self.navigation_frame, text="", image=self.logo_image,
                                                             compound="left", font=customtkinter.CTkFont(size=15, weight="bold"))
        self.navigation_frame_label.grid(row=0, column=0, padx=50, pady=20)

        self.frame_1_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Savaitinės gamintojų ataskaitos",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), font=("Calibri", 14),
                                                   image=self.frame_1_image, anchor="w", command=self.frame_1_button_event)
        self.frame_1_button.grid(row=1, column=0, sticky="ew")

        self.frame_2_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="SPEC savaitinė likučių ataskaita",
                                                      fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), font=("Calibri", 14),
                                                      image=self.frame_2_image, anchor="w", command=self.frame_2_button_event)
        self.frame_2_button.grid(row=2, column=0, sticky="ew")

        self.frame_3_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="SPEC mėnesinė ataskaita",
                                                      fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), font=("Calibri", 14),
                                                      image=self.frame_3_image, anchor="w", command=self.frame_3_button_event)
        self.frame_3_button.grid(row=3, column=0, sticky="ew")

        self.appearance_mode_menu = customtkinter.CTkOptionMenu(self.navigation_frame, values=["Light", "Dark", "System"],
                                                                command=self.change_appearance_mode_event)
        self.appearance_mode_menu.grid(row=6, column=0, padx=50, pady=20, sticky="s")

        # create first frame

        self.frame_1 = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.frame_1.grid_columnconfigure(1, weight=1)

        self.frame_label_1 = customtkinter.CTkLabel(master=self.frame_1, text="Savaitinės gamintojų ataskaitos", font=("Calibri", 20))
        self.frame_label_1.grid(row=0, column=0, columnspan=2, padx=20, pady=10)

        self.browse_label_1 = customtkinter.CTkLabel(master=self.frame_1, text="Ataskaitų eksportavimo vieta:", font=("Calibri", 16))
        self.browse_label_1.grid(row=1, column=0, columnspan=2, padx=0, pady=2)

        self.browse_entry_1 = customtkinter.CTkEntry(master=self.frame_1, width=300, font=("Calibri", 12), state="disabled", textvariable=default_export_text)
        self.browse_entry_1.grid(row=2, column=0, padx=(40, 0), pady=5, sticky="e")

        self.browse_button_1 = customtkinter.CTkButton(master=self.frame_1, text="Naršyti...", command=lambda: self.browse(default_export_text), width=60, font=("Calibri", 16))
        self.browse_button_1.grid(row=2, column=1, padx=(5, 0), pady=5, sticky="w")

        self.settings_label_1 = customtkinter.CTkLabel(master=self.frame_1, text="Nustatymai:", font=("Calibri", 16))
        self.settings_label_1.grid(row=3, column=0, columnspan=2, padx=0, pady=(10, 2))

        self.manufacturer_button_1 = customtkinter.CTkButton(master=self.frame_1, text="Gamintojų sąrašas", command=self.manufacturer_edit, width=200, font=("Calibri", 16))
        self.manufacturer_button_1.grid(row=4, column=0, columnspan=2, padx=(0, 0), pady=5)

        self.settings_button_1 = customtkinter.CTkButton(master=self.frame_1, text="Eksportavimo nustatymai", command=self.export_settings, width=200, font=("Calibri", 16))
        self.settings_button_1.grid(row=5, column=0, columnspan=2, padx=(0, 0), pady=5)

        self.spec_checked = tk.IntVar()
        self.spec_checkbox_1 = customtkinter.CTkCheckBox(master=self.frame_1, text="Pabaigus - Pradėti SPEC likučius", width=250, font=("Calibri", 16), variable=self.spec_checked)
        self.spec_checkbox_1.grid(row=6, column=0, columnspan=2, padx=(0, 0), pady=(20, 5))

        self.close_checked_1 = tk.IntVar()
        self.close_checkbox_1 = customtkinter.CTkCheckBox(master=self.frame_1, text="Pabaigus - Išjungti įmonės programa", width=250, font=("Calibri", 16), variable=self.close_checked_1)
        self.close_checkbox_1.grid(row=7, column=0, columnspan=2, padx=(0, 0), pady=5)

        self.login_frame_1 = customtkinter.CTkFrame(master=self.frame_1, width=400, height=300, fg_color="transparent")
        self.login_frame_1.grid(row=8, column=0, rowspan=3, columnspan=2, padx=20, pady=10)

        self.login_label_1 = customtkinter.CTkLabel(master=self.login_frame_1, text="Prisijungimo duomenys:", font=("Calibri", 16))
        self.login_label_1.grid(row=0, column=0, columnspan=2, padx=0, pady=5)

        self.user_name_1_label = customtkinter.CTkLabel(master=self.login_frame_1, width=100, text="Vartotojas:", font=("Calibri", 16))
        self.user_name_1_label.grid(row=1, column=0, padx=0, pady=3)
    
        self.user_name_1 = tk.StringVar()
        self.user_name_1_entry = customtkinter.CTkEntry(master=self.login_frame_1, width=150, font=("Calibri", 16), textvariable=self.user_name_1)
        self.user_name_1_entry.grid(row=1, column=1, padx=0, pady=3)

        self.user_password_1_label = customtkinter.CTkLabel(master=self.login_frame_1, width=100, text="Slaptažodis:", font=("Calibri", 16))
        self.user_password_1_label.grid(row=2, column=0, padx=0, pady=3)
        
        self.user_password_1 = tk.StringVar()
        self.user_password_1_entry = customtkinter.CTkEntry(master=self.login_frame_1, width=150, font=("Calibri", 16), show="●", textvariable=self.user_password_1)
        self.user_password_1_entry.bind("<Return>", (lambda event: threading.Thread(target=self.initiate_weekly_report_module).start()))
        self.user_password_1_entry.grid(row=2, column=1, padx=0, pady=3)

        self.start_button_1 = customtkinter.CTkButton(master=self.frame_1, text="Pradėti", width=200, font=("Calibri", 25), fg_color="green", hover_color="darkgreen", command=lambda: threading.Thread(target=self.initiate_weekly_report_module).start())
        self.start_button_1.grid(row=11, column=0, columnspan=2, padx=(0, 0), pady=15)

        # create second frame

        self.frame_2 = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.frame_2.grid_columnconfigure(1, weight=1)

        self.frame_label_2 = customtkinter.CTkLabel(master=self.frame_2, text="SPEC savaitinė likučių ataskaita", font=("Calibri", 20))
        self.frame_label_2.grid(row=0, column=0, columnspan=3, padx=20, pady=10)

        self.browse_label_2 = customtkinter.CTkLabel(master=self.frame_2, text="Ataskaitų eksportavimo vieta:", font=("Calibri", 16))
        self.browse_label_2.grid(row=1, column=0, columnspan=3, padx=0, pady=2)

        self.browse_entry_2 = customtkinter.CTkEntry(master=self.frame_2, width=300, font=("Calibri", 12), state="disabled", textvariable=default_export_text)
        self.browse_entry_2.grid(row=2, column=0, padx=(40, 0), pady=5, sticky="e")

        self.browse_button_2 = customtkinter.CTkButton(master=self.frame_2, text="Naršyti...", command=lambda: self.browse(default_export_text), width=60, font=("Calibri", 16))
        self.browse_button_2.grid(row=2, column=1, padx=(5, 0), pady=5, sticky="w")

        self.settings_label_2_main = customtkinter.CTkLabel(master=self.frame_2, text="Nustatymai:", font=("Calibri", 16))
        self.settings_label_2_main.grid(row=3, column=0, columnspan=3, padx=0, pady=(10, 2))

        self.settings_button_2 = customtkinter.CTkButton(master=self.frame_2, text="SPEC komponentai", command=self.spec_component_settings, width=200, font=("Calibri", 16))
        self.settings_button_2.grid(row=4, column=0, columnspan=3, padx=(0, 0), pady=5)

        self.close_checked_2 = tk.IntVar()
        self.close_checkbox_2 = customtkinter.CTkCheckBox(master=self.frame_2, text="Pabaigus - Išjungti įmonės programa", width=250, font=("Calibri", 16), variable=self.close_checked_2)
        self.close_checkbox_2.grid(row=5, column=0, columnspan=2, padx=(0, 0), pady=(15, 0))

        self.login_frame_2 = customtkinter.CTkFrame(master=self.frame_2, width=400, height=300, fg_color="transparent")
        self.login_frame_2.grid(row=6, column=0, rowspan=3, columnspan=2, padx=20, pady=10)

        self.login_label_2 = customtkinter.CTkLabel(master=self.login_frame_2, text="Prisijungimo duomenys:", font=("Calibri", 16))
        self.login_label_2.grid(row=0, column=0, columnspan=2, padx=0, pady=5)

        self.user_name_2_label = customtkinter.CTkLabel(master=self.login_frame_2, width=100, text="Vartotojas:", font=("Calibri", 16))
        self.user_name_2_label.grid(row=1, column=0, padx=0, pady=3)
    
        self.user_name_2 = tk.StringVar()
        self.user_name_2_entry = customtkinter.CTkEntry(master=self.login_frame_2, width=150, font=("Calibri", 16), textvariable=self.user_name_2)
        self.user_name_2_entry.grid(row=1, column=1, padx=0, pady=3)

        self.user_password_2_label = customtkinter.CTkLabel(master=self.login_frame_2, width=100, text="Slaptažodis:", font=("Calibri", 16))
        self.user_password_2_label.grid(row=2, column=0, padx=0, pady=3)
        
        self.user_password_2 = tk.StringVar()
        self.user_password_2_entry = customtkinter.CTkEntry(master=self.login_frame_2, width=150, font=("Calibri", 16), show="●", textvariable=self.user_password_2)
        self.user_password_2_entry.bind("<Return>", (lambda event: threading.Thread(target=self.initiate_spec_weekly_report_module).start()))
        self.user_password_2_entry.grid(row=2, column=1, padx=0, pady=3)

        self.start_button_2 = customtkinter.CTkButton(master=self.frame_2, text="Pradėti", width=200, font=("Calibri", 25), fg_color="green", hover_color="darkgreen", command=lambda: threading.Thread(target=self.initiate_spec_weekly_report_module).start())
        self.start_button_2.grid(row=10, column=0, columnspan=2, padx=(0, 0), pady=15)

        # create third frame

        self.frame_3 = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.frame_3.grid_columnconfigure(1, weight=1)

        self.frame_label_3 = customtkinter.CTkLabel(master=self.frame_3, text="SPEC mėnesinė ataskaita", font=("Calibri", 20))
        self.frame_label_3.grid(row=0, column=0, columnspan=3, padx=20, pady=10)

        self.browse_label_3 = customtkinter.CTkLabel(master=self.frame_3, text="Ataskaitų eksportavimo vieta:", font=("Calibri", 16))
        self.browse_label_3.grid(row=1, column=0, columnspan=3, padx=0, pady=2)

        self.browse_entry_3 = customtkinter.CTkEntry(master=self.frame_3, width=300, font=("Calibri", 12), state="disabled", textvariable=default_export_text)
        self.browse_entry_3.grid(row=2, column=0, padx=(40, 0), pady=5, sticky="e")

        self.browse_button_3 = customtkinter.CTkButton(master=self.frame_3, text="Naršyti...", command=lambda: self.browse(default_export_text), width=60, font=("Calibri", 16))
        self.browse_button_3.grid(row=2, column=1, padx=(5, 0), pady=5, sticky="w")

        self.settings_label_3_main = customtkinter.CTkLabel(master=self.frame_3, text="Nustatymai:", font=("Calibri", 16))
        self.settings_label_3_main.grid(row=3, column=0, columnspan=3, padx=0, pady=(10, 2))

        self.settings_button_3 = customtkinter.CTkButton(master=self.frame_3, text="SPEC komponentai", command=self.spec_component_settings, width=200, font=("Calibri", 16))
        self.settings_button_3.grid(row=4, column=0, columnspan=3, padx=(0, 0), pady=5)

        self.close_checked_3 = tk.IntVar()
        self.close_checkbox_3 = customtkinter.CTkCheckBox(master=self.frame_3, text="Pabaigus - Išjungti įmonės programa", width=250, font=("Calibri", 16), variable=self.close_checked_3)
        self.close_checkbox_3.grid(row=5, column=0, columnspan=2, padx=(0, 0), pady=(15, 0))

        self.login_frame_3 = customtkinter.CTkFrame(master=self.frame_3, width=400, height=300, fg_color="transparent")
        self.login_frame_3.grid(row=6, column=0, rowspan=3, columnspan=2, padx=20, pady=10)

        self.login_label_3 = customtkinter.CTkLabel(master=self.login_frame_3, text="Prisijungimo duomenys:", font=("Calibri", 16))
        self.login_label_3.grid(row=0, column=0, columnspan=2, padx=0, pady=5)

        self.user_name_3_label = customtkinter.CTkLabel(master=self.login_frame_3, width=100, text="Vartotojas:", font=("Calibri", 16))
        self.user_name_3_label.grid(row=1, column=0, padx=0, pady=3)
    
        self.user_name_3 = tk.StringVar()
        self.user_name_3_entry = customtkinter.CTkEntry(master=self.login_frame_3, width=150, font=("Calibri", 16), textvariable=self.user_name_3)
        self.user_name_3_entry.grid(row=1, column=1, padx=0, pady=3)

        self.user_password_3_label = customtkinter.CTkLabel(master=self.login_frame_3, width=100, text="Slaptažodis:", font=("Calibri", 16))
        self.user_password_3_label.grid(row=2, column=0, padx=0, pady=3)
        
        self.user_password_3 = tk.StringVar()
        self.user_password_3_entry = customtkinter.CTkEntry(master=self.login_frame_3, width=150, font=("Calibri", 16), show="●", textvariable=self.user_password_3)
        self.user_password_3_entry.bind("<Return>", (lambda event: threading.Thread(target=self.initiate_spec_monthly_report_module).start()))
        self.user_password_3_entry.grid(row=2, column=1, padx=0, pady=3)

        self.start_button_3 = customtkinter.CTkButton(master=self.frame_3, text="Pradėti", width=200, font=("Calibri", 25), fg_color="green", hover_color="darkgreen", command=lambda: threading.Thread(target=self.initiate_spec_monthly_report_module).start())
        self.start_button_3.grid(row=10, column=0, columnspan=2, padx=(0, 0), pady=15)

        # select default frame
        self.select_frame_by_name("frame_1")

    def select_frame_by_name(self, name):
        # set button color for selected button
        self.frame_1_button.configure(fg_color=("gray75", "gray25") if name == "frame_1" else "transparent")
        self.frame_2_button.configure(fg_color=("gray75", "gray25") if name == "frame_2" else "transparent")
        self.frame_3_button.configure(fg_color=("gray75", "gray25") if name == "frame_3" else "transparent")

        # show selected frame
        if name == "frame_1":
            self.frame_1.grid(row=0, column=1, sticky="nsew")
        else:
            self.frame_1.grid_forget()
        if name == "frame_2":
            self.frame_2.grid(row=0, column=1, sticky="nsew")
        else:
            self.frame_2.grid_forget()
        if name == "frame_3":
            self.frame_3.grid(row=0, column=1, sticky="nsew")
        else:
            self.frame_3.grid_forget()

    def frame_1_button_event(self):
        self.select_frame_by_name("frame_1")

    def frame_2_button_event(self):
        self.select_frame_by_name("frame_2")

    def frame_3_button_event(self):
        self.select_frame_by_name("frame_3")

    def change_appearance_mode_event(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)
    
    def browse(self, default_export_text):
        folder_selected = filedialog.askdirectory(title="Pasirinkite aplanka")
        if folder_selected == "":
            folder_selected = default_export_text.get()
        default_export_text.set(folder_selected)
        data = {
            "directories": {
                            "report_directory": folder_selected
                        }
        }
        with open("data/DirectoryData.json", "w") as f:
            json.dump(data, f, indent=4)
    
    def manufacturer_edit(self):
        self.frame_label_1.grid_forget()
        self.browse_label_1.grid_forget()
        self.browse_entry_1.grid_forget()
        self.browse_button_1.grid_forget()
        self.settings_label_1.grid_forget()
        self.manufacturer_button_1.grid_forget()
        self.settings_button_1.grid_forget()
        self.spec_checkbox_1.grid_forget()
        self.close_checkbox_1.grid_forget()
        self.start_button_1.grid_forget()
        self.login_frame_1.grid_forget()
        self.login_label_1.grid_forget()
        self.user_name_1_label.grid_forget()
        self.user_name_1_entry.grid_forget()
        self.user_password_1_label.grid_forget()
        self.user_password_1_entry.grid_forget()
        
        self.manufacturer_list = []
        f = open("data/ManufacturerData.json",)
        f_data = json.load(f)
        for manufacturer in f_data["manufacturers"]:
            self.manufacturer_list.append(manufacturer)
        f.close()
        manufacturer_entry_text = tk.StringVar(value="")
        
        self.frame_label_1 = customtkinter.CTkLabel(master=self.frame_1, text="Gamintojų sąrašas", font=("Calibri", 20))
        self.frame_label_1.grid(row=0, column=0, columnspan=2, padx=20, pady=10)

        self.manufacturer_frame = customtkinter.CTkScrollableFrame(master=self.frame_1, width=175, height=300)
        self.manufacturer_frame.grid(row=1, column=0, rowspan=4, padx=20, pady=10)
        self.manufacturer = []
        self.manufacturer_del = []
        for row_idx, m_name in enumerate(self.manufacturer_list):
            self.manufacturer.append(customtkinter.CTkLabel(master=self.manufacturer_frame, text=m_name, font=("Calibri", 14), width=100))
            self.manufacturer[row_idx].grid(row=row_idx, column=1, padx=(0, 0), pady=2)
            self.manufacturer_del.append(customtkinter.CTkButton(master=self.manufacturer_frame, text="Ištrinti", width=50, font=("Calibri", 14), fg_color="darkgrey", text_color="black", hover_color="firebrick",
                                                       command=lambda row_idx=row_idx: self.delete_manufacturer(self.manufacturer[row_idx], self.manufacturer_del[row_idx], row_idx)))
            self.manufacturer_del[row_idx].grid(row=row_idx, column=2, padx=(10, 0), pady=2, sticky="ne")

        self.new_manufacturer_label = customtkinter.CTkLabel(master=self.frame_1, text="Pridėti nauja gamintoja:", font=("Calibri", 16))
        self.new_manufacturer_label.grid(row=2, column=1)
        self.new_manufacturer_entry = customtkinter.CTkEntry(master=self.frame_1, width=175, font=("Calibri", 16), textvariable=manufacturer_entry_text, placeholder_text="Gamintojas", border_color="black")
        self.new_manufacturer_entry.grid(row=2, column=1, sticky="s", pady=5)
        self.new_manufacturer_button = customtkinter.CTkButton(master=self.frame_1, text="Pridėti", command=lambda: self.add_manufacturer(manufacturer_entry_text), width=100, font=("Calibri", 16))
        self.new_manufacturer_button.grid(row=3, column=1, sticky="n", pady=5)
        self.error_label_1 = customtkinter.CTkLabel(master=self.frame_1, text="Pavadinimas negali būti\ntrumpesnis nei 3 \nsimboliai, ir ne ilgesnis\nnei 15 simbolių", font=("Calibri", 16), text_color="red")
        self.error_label_1.grid(row=3, column=1, sticky="s", pady=5)
        self.error_label_1.grid_forget()

        self.save_button_1 = customtkinter.CTkButton(master=self.frame_1, text="Išsaugoti", width=100, font=("Calibri", 16), command=self.save_manufacturer_info)
        self.save_button_1.grid(row=5, column=0, padx=(125, 5), pady=5, sticky="ne")
        self.back_button_1 = customtkinter.CTkButton(master=self.frame_1, text="<< Grįžti", width=100, font=("Calibri", 16), fg_color="gainsboro", text_color="black", hover_color="darkgrey", command=self.go_back_manufacturer)
        self.back_button_1.grid(row=5, column=1, padx=(5, 125), pady=5, sticky="nw")

    
    def add_manufacturer(self, manufacturer_entry_text):
        manufacturer_name = self.new_manufacturer_entry.get()
        if len(manufacturer_name) > 2 and len(manufacturer_name) < 15:
            self.manufacturer_list.append(manufacturer_name)
            row_idx = len(self.manufacturer)
            self.new_manufacturer_entry.configure(border_color="black")
            manufacturer_entry_text.set("")
            self.error_label_1.grid_forget()

            self.manufacturer.append(customtkinter.CTkLabel(master=self.manufacturer_frame, text=manufacturer_name, font=("Calibri", 14), width=100))
            self.manufacturer[row_idx].grid(row=row_idx, column=1, padx=(0, 0), pady=2)
            self.manufacturer_del.append(customtkinter.CTkButton(master=self.manufacturer_frame, text="Ištrinti", width=50, font=("Calibri", 14), fg_color="darkgrey", text_color="black", hover_color="firebrick",
                                                       command=lambda row_idx=row_idx: self.delete_manufacturer(self.manufacturer[row_idx], self.manufacturer_del[row_idx], row_idx)))
            self.manufacturer_del[row_idx].grid(row=row_idx, column=2, padx=(10, 0), pady=2, sticky="ne")
        else:
            self.new_manufacturer_entry.configure(border_color="red")
            self.error_label_1.grid(row=3, column=1, sticky="s", pady=5)

    def delete_manufacturer(self, specific_manufacturer, specific_manufacturer_del, row_index):
        specific_manufacturer_del.destroy()
        specific_manufacturer.destroy()
        self.manufacturer_list[row_index] = "TO_BE_DELETED"
    
    def save_manufacturer_info(self):
        self.manufacturer_list = [value for value in self.manufacturer_list if value != "TO_BE_DELETED"]
        data = {
            "manufacturers": self.manufacturer_list
        }
        with open("data/ManufacturerData.json", "w") as f:
            json.dump(data, f, indent=4)
        
        self.manufacturer_frame.grid_forget()
        self.new_manufacturer_label.grid_forget()
        self.new_manufacturer_entry.grid_forget()
        self.new_manufacturer_button.grid_forget()
        self.error_label_1.grid_forget()
        self.save_button_1.grid_forget()
        self.back_button_1.grid_forget()
        self.manufacturer_frame = customtkinter.CTkScrollableFrame(master=self.frame_1, width=175,  height=300)
        self.manufacturer_frame.grid(row=1, column=0, rowspan=4, padx=20, pady=10)
        self.manufacturer = []
        self.manufacturer_del = []
        for row_idx, m_name in enumerate(self.manufacturer_list):
            self.manufacturer.append(customtkinter.CTkLabel(master=self.manufacturer_frame, text=m_name, font=("Calibri", 14), width=100))
            self.manufacturer[row_idx].grid(row=row_idx, column=1, padx=(0, 0), pady=2)
            self.manufacturer_del.append(customtkinter.CTkButton(master=self.manufacturer_frame, text="Ištrinti", width=50, font=("Calibri", 14), fg_color="darkgrey", text_color="black", hover_color="firebrick",
                                                       command=lambda row_idx=row_idx: self.delete_manufacturer(self.manufacturer[row_idx], self.manufacturer_del[row_idx], row_idx)))
            self.manufacturer_del[row_idx].grid(row=row_idx, column=2, padx=(10, 0), pady=2, sticky="ne")

        manufacturer_entry_text = tk.StringVar(value="")
        self.new_manufacturer_label = customtkinter.CTkLabel(master=self.frame_1, text="Pridėti nauja gamintoja:", font=("Calibri", 16))
        self.new_manufacturer_label.grid(row=2, column=1)
        self.new_manufacturer_entry = customtkinter.CTkEntry(master=self.frame_1, width=175, font=("Calibri", 16), textvariable=manufacturer_entry_text, placeholder_text="Gamintojas", border_color="black")
        self.new_manufacturer_entry.grid(row=2, column=1, sticky="s", pady=5)
        self.new_manufacturer_button = customtkinter.CTkButton(master=self.frame_1, text="Pridėti", command=lambda: self.add_manufacturer(manufacturer_entry_text), width=100, font=("Calibri", 16))
        self.new_manufacturer_button.grid(row=3, column=1, sticky="n", pady=5)

        self.error_label_1 = customtkinter.CTkLabel(master=self.frame_1, text="Pavadinimas negali būti\ntrumpesnis nei 3 \nsimboliai, ir ne ilgesnis\nnei 15 simbolių", font=("Calibri", 16), text_color="red")
        self.error_label_1.grid(row=3, column=1, sticky="s", pady=5)
        self.error_label_1.grid_forget()

        self.save_button_1 = customtkinter.CTkButton(master=self.frame_1, text="Išsaugoti", width=100, font=("Calibri", 16), command=self.save_manufacturer_info)
        self.save_button_1.grid(row=5, column=0, padx=(125, 5), pady=5, sticky="ne")
        self.back_button_1 = customtkinter.CTkButton(master=self.frame_1, text="<< Grįžti", width=100, font=("Calibri", 16), fg_color="gainsboro", text_color="black", hover_color="darkgrey", command=self.go_back_manufacturer)
        self.back_button_1.grid(row=5, column=1, padx=(5, 125), pady=5, sticky="nw")
    
    def go_back_manufacturer(self):
        self.frame_label_1.grid_forget()
        self.manufacturer_frame.grid_forget()
        self.new_manufacturer_label.grid_forget()
        self.new_manufacturer_entry.grid_forget()
        self.new_manufacturer_button.grid_forget()
        self.error_label_1.grid_forget()
        self.save_button_1.grid_forget()
        self.back_button_1.grid_forget()

        self.frame_label_1.grid(row=0, column=0, columnspan=2, padx=20, pady=10)
        self.browse_label_1.grid(row=1, column=0, columnspan=2, padx=0, pady=2)
        self.browse_entry_1.grid(row=2, column=0, padx=(40, 0), pady=5, sticky="e")
        self.browse_button_1.grid(row=2, column=1, padx=(5, 0), pady=5, sticky="w")
        self.settings_label_1.grid(row=3, column=0, columnspan=2, padx=0, pady=(10, 2))
        self.manufacturer_button_1.grid(row=4, column=0, columnspan=2, padx=(0, 0), pady=5)
        self.settings_button_1.grid(row=5, column=0, columnspan=2, padx=(0, 0), pady=5)
        self.spec_checkbox_1.grid(row=6, column=0, columnspan=2, padx=(0, 0), pady=(20, 5))
        self.close_checkbox_1.grid(row=7, column=0, columnspan=2, padx=(0, 0), pady=5)
        self.start_button_1.grid(row=11, column=0, columnspan=2, padx=(0, 0), pady=15)
        self.login_frame_1.grid(row=8, column=0, rowspan=3, columnspan=2, padx=20, pady=10)
        self.login_label_1.grid(row=0, column=0, columnspan=2, padx=0, pady=5)
        self.user_name_1_label.grid(row=1, column=0, padx=0, pady=3)
        self.user_name_1_entry.grid(row=1, column=1, padx=0, pady=3)
        self.user_password_1_label.grid(row=2, column=0, padx=0, pady=3)
        self.user_password_1_entry.grid(row=2, column=1, padx=0, pady=3)


    def export_settings(self):
        self.frame_label_1.grid_forget()
        self.browse_label_1.grid_forget()
        self.browse_entry_1.grid_forget()
        self.browse_button_1.grid_forget()
        self.settings_label_1.grid_forget()
        self.manufacturer_button_1.grid_forget()
        self.settings_button_1.grid_forget()
        self.spec_checkbox_1.grid_forget()
        self.close_checkbox_1.grid_forget()
        self.start_button_1.grid_forget()
        self.login_frame_1.grid_forget()
        self.login_label_1.grid_forget()
        self.user_name_1_label.grid_forget()
        self.user_name_1_entry.grid_forget()
        self.user_password_1_label.grid_forget()
        self.user_password_1_entry.grid_forget()

        self.export_list = []
        self.filter_list = []
        f = open("data/ExportSettings.json",)
        f_data = json.load(f)
        for export in f_data:
            temp_list = []
            self.export_list.append(export)
            for filter in f_data[export]:
                temp_list.append(filter)
            self.filter_list.append(temp_list)
        f.close()

        self.frame_label_1 = customtkinter.CTkLabel(master=self.frame_1, text="Eksportavimo nustatymai", font=("Calibri", 20))
        self.frame_label_1.grid(row=0, column=0, columnspan=2, padx=20, pady=10)

        self.report_frame = customtkinter.CTkScrollableFrame(master=self.frame_1, width=150, height=200)
        self.report_frame.grid(row=1, column=0, rowspan=3, padx=20, pady=10)

        self.report = []
        self.report_edit = []
        for row_idx, r_name in enumerate(self.export_list):
            self.report.append(customtkinter.CTkLabel(master=self.report_frame, text=r_name, font=("Calibri", 14), width=105))
            self.report[row_idx].grid(row=row_idx, column=0, padx=(0, 0), pady=2)
            self.report_edit.append(customtkinter.CTkButton(master=self.report_frame, text="\u270e", width=30, font=("Arial Unicode MS", 18), fg_color="darkgrey", text_color="black", hover_color="grey",
                                                            command=lambda row_idx=row_idx: self.edit_report_filter(self.report[row_idx], self.report_edit[row_idx], self.filter_list[row_idx], row_idx)))
            self.report_edit[row_idx].grid(row=row_idx, column=1, padx=(10, 0), pady=2, sticky="ne")

        report_entry_text = tk.StringVar(value="")
        self.new_report_label = customtkinter.CTkLabel(master=self.frame_1, text="Pridėti nauja ataskaita:", font=("Calibri", 16))
        self.new_report_label.grid(row=2, column=1, sticky="n", pady=5)
        self.new_report_entry = customtkinter.CTkEntry(master=self.frame_1, width=175, font=("Calibri", 16), textvariable=report_entry_text, placeholder_text="Ataskaita", border_color="black")
        self.new_report_entry.grid(row=2, column=1, sticky="s", pady=5)
        self.new_report_button = customtkinter.CTkButton(master=self.frame_1, text="Pridėti", width=100, font=("Calibri", 16), command=lambda: self.add_report(report_entry_text))
        self.new_report_button.grid(row=3, column=1, sticky="n", pady=5)
        self.error_label_1 = customtkinter.CTkLabel(master=self.frame_1, text="Pavadinimas negali būti\ntrumpesnis nei 3 \nsimboliai, ir ne ilgesnis\nnei 15 simbolių", font=("Calibri", 16), text_color="red")
        self.error_label_1.grid(row=3, column=1, sticky="s", pady=5)
        self.error_label_1.grid_forget()

        self.report_edit_frame = customtkinter.CTkFrame(master=self.frame_1, width=300,  height=100)
        self.report_edit_label = customtkinter.CTkLabel(master=self.report_edit_frame, text="Filtrų redagavimas", font=("Calibri", 18))
        self.group_edit_label = customtkinter.CTkLabel(master=self.report_edit_frame, text="Grupė", font=("Calibri", 16))
        self.sku_edit_label = customtkinter.CTkLabel(master=self.report_edit_frame, text="Kodas", font=("Calibri", 16))
        self.name_edit_label = customtkinter.CTkLabel(master=self.report_edit_frame, text="Pavadinimas", font=("Calibri", 16))
        self.group_edit_entry = customtkinter.CTkEntry(master=self.report_edit_frame, width=100, font=("Calibri", 16), textvariable="", border_color="black")
        self.sku_edit_entry = customtkinter.CTkEntry(master=self.report_edit_frame, width=100, font=("Calibri", 16), textvariable="", border_color="black")
        self.name_edit_entry = customtkinter.CTkEntry(master=self.report_edit_frame, width=100, font=("Calibri", 16), textvariable="", border_color="black")
        self.report_edit_save = customtkinter.CTkButton(master=self.report_edit_frame, text="Pakeisti filtro nustatymus", width=150, font=("Calibri", 16))
        self.report_edit_delete = customtkinter.CTkButton(master=self.report_edit_frame, text="Ištrinti ataskaita", width=100, font=("Calibri", 16), fg_color="firebrick", hover_color="darkred")

        self.save_button_1 = customtkinter.CTkButton(master=self.frame_1, text="Išsaugoti", width=100, font=("Calibri", 16), command=self.save_export_settings)
        self.save_button_1.grid(row=5, column=0, padx=(125, 5), pady=10, sticky="ne")
        self.back_button_1 = customtkinter.CTkButton(master=self.frame_1, text="<< Grįžti", width=100, font=("Calibri", 16), fg_color="gainsboro", text_color="black", hover_color="darkgrey",
                                                     command=self.go_back_export_settings)
        self.back_button_1.grid(row=5, column=1, padx=(5, 125), pady=10, sticky="nw")

    
    def add_report(self, report_entry_text):
        report_name = self.new_report_entry.get()
        if len(report_name) > 2 and len(report_name) < 15:
            self.export_list.append(report_name)
            self.filter_list.append(["", "", ""])
            row_idx = len(self.report)
            self.new_report_entry.configure(border_color="black")
            report_entry_text.set("")
            self.error_label_1.grid_forget()

            self.report.append(customtkinter.CTkLabel(master=self.report_frame, text=report_name, font=("Calibri", 14), width=105))
            self.report[row_idx].grid(row=row_idx, column=0, padx=(0, 0), pady=2)
            self.report_edit.append(customtkinter.CTkButton(master=self.report_frame, text="\u270e", width=30, font=("Arial Unicode MS", 18), fg_color="darkgrey", text_color="black", hover_color="grey",
                                                            command=lambda row_idx=row_idx: self.edit_report_filter(self.report[row_idx], self.report_edit[row_idx], self.filter_list[row_idx], row_idx)))
            self.report_edit[row_idx].grid(row=row_idx, column=1, padx=(10, 0), pady=2, sticky="ne")
        else:
            self.new_report_entry.configure(border_color="red")
            self.error_label_1.grid(row=3, column=1, sticky="s", pady=5)
    
    def edit_report_filter(self, specific_report, specific_edit, filters, row_index):
        main_label_text = specific_report.cget("text") + " filtrų nustatymai"
        self.report_edit_label.configure(text=main_label_text)
        group_edit_entry_text = tk.StringVar(value=filters[0])
        sku_edit_entry_text = tk.StringVar(value=filters[1])
        name_edit_entry_text = tk.StringVar(value=filters[2])
        self.group_edit_entry.configure(textvariable=group_edit_entry_text)
        self.sku_edit_entry.configure(textvariable=sku_edit_entry_text)
        self.name_edit_entry.configure(textvariable=name_edit_entry_text)
        self.report_edit_save.configure(command=lambda: self.leave_and_save_filter_settings(row_index))
        self.report_edit_delete.configure(command=lambda: self.delete_report(specific_report, specific_edit, row_index))
        
        self.report_edit_frame.grid(row=4, column=0, columnspan=2, padx=20, pady=10)
        self.report_edit_label.grid(row=0, column=0, columnspan=3, padx=20, pady=(5, 0))
        self.group_edit_label.grid(row=1, column=0, padx=15, pady=5)
        self.sku_edit_label.grid(row=1, column=1, padx=15, pady=5)
        self.name_edit_label.grid(row=1, column=2, padx=15, pady=5)
        self.group_edit_entry.grid(row=2, column=0, padx=15, pady=(0, 10))
        self.sku_edit_entry.grid(row=2, column=1, padx=15, pady=(0, 10))
        self.name_edit_entry.grid(row=2, column=2, padx=15, pady=(0, 10))
        self.report_edit_save.grid(row=3, column=0, columnspan=3, sticky="nw", padx=20, pady=(5,10))
        self.report_edit_delete.grid(row=3, column=0, columnspan=3, sticky="ne", padx=20, pady=(5,10))

    def delete_report(self, specific_report, specific_edit, row_index):
        specific_report.destroy()
        specific_edit.destroy()
        self.export_list[row_index] = "TO_BE_DELETED"
        self.filter_list[row_index][0] = "TO_BE_DELETED"
        self.filter_list[row_index][1] = "TO_BE_DELETED"
        self.filter_list[row_index][2] = "TO_BE_DELETED"

        self.report_edit_frame.grid_forget()
        self.report_edit_label.grid_forget()
        self.group_edit_label.grid_forget()
        self.sku_edit_label.grid_forget()
        self.name_edit_label.grid_forget()
        self.group_edit_entry.grid_forget()
        self.sku_edit_entry.grid_forget()
        self.name_edit_entry.grid_forget()
        self.report_edit_save.grid_forget()
        self.report_edit_delete.grid_forget()

    def leave_and_save_filter_settings(self, row_index):
        self.filter_list[row_index][0] = self.group_edit_entry.get()
        self.filter_list[row_index][1] = self.sku_edit_entry.get()
        self.filter_list[row_index][2] = self.name_edit_entry.get()

        self.report_edit_frame.grid_forget()
        self.report_edit_label.grid_forget()
        self.group_edit_label.grid_forget()
        self.sku_edit_label.grid_forget()
        self.name_edit_label.grid_forget()
        self.group_edit_entry.grid_forget()
        self.sku_edit_entry.grid_forget()
        self.name_edit_entry.grid_forget()
        self.report_edit_save.grid_forget()
        self.report_edit_delete.grid_forget()
        print(self.filter_list)

    def go_back_export_settings(self):
        self.frame_label_1.grid_forget()
        self.report_frame.grid_forget()
        self.new_report_label.grid_forget()
        self.new_report_entry.grid_forget()
        self.new_report_button.grid_forget()
        self.error_label_1.grid_forget()
        self.save_button_1.grid_forget()
        self.back_button_1.grid_forget()
        self.report_edit_frame.grid_forget()
        self.report_edit_label.grid_forget()
        self.group_edit_label.grid_forget()
        self.sku_edit_label.grid_forget()
        self.name_edit_label.grid_forget()
        self.group_edit_entry.grid_forget()
        self.sku_edit_entry.grid_forget()
        self.name_edit_entry.grid_forget()
        self.report_edit_save.grid_forget()
        self.report_edit_delete.grid_forget()
        
        self.frame_label_1.grid(row=0, column=0, columnspan=2, padx=20, pady=10)
        self.browse_label_1.grid(row=1, column=0, columnspan=2, padx=0, pady=2)
        self.browse_entry_1.grid(row=2, column=0, padx=(40, 0), pady=5, sticky="e")
        self.browse_button_1.grid(row=2, column=1, padx=(5, 0), pady=5, sticky="w")
        self.settings_label_1.grid(row=3, column=0, columnspan=2, padx=0, pady=(10, 2))
        self.manufacturer_button_1.grid(row=4, column=0, columnspan=2, padx=(0, 0), pady=5)
        self.settings_button_1.grid(row=5, column=0, columnspan=2, padx=(0, 0), pady=5)
        self.spec_checkbox_1.grid(row=6, column=0, columnspan=2, padx=(0, 0), pady=(20, 5))
        self.close_checkbox_1.grid(row=7, column=0, columnspan=2, padx=(0, 0), pady=5)
        self.start_button_1.grid(row=11, column=0, columnspan=2, padx=(0, 0), pady=15)
        self.login_frame_1.grid(row=8, column=0, rowspan=3, columnspan=2, padx=20, pady=10)
        self.login_label_1.grid(row=0, column=0, columnspan=2, padx=0, pady=5)
        self.user_name_1_label.grid(row=1, column=0, padx=0, pady=3)
        self.user_name_1_entry.grid(row=1, column=1, padx=0, pady=3)
        self.user_password_1_label.grid(row=2, column=0, padx=0, pady=3)
        self.user_password_1_entry.grid(row=2, column=1, padx=0, pady=3)
    
    def save_export_settings(self):
        new_export_list = []
        new_filter_list = []
        for idx, report in enumerate(self.export_list):
            if report != "TO_BE_DELETED":
                temp_list = []
                new_export_list.append(report)
                for filter in self.filter_list[idx]:
                    temp_list.append(filter)
                new_filter_list.append(temp_list)
        
        data = {}
        for idx, report in enumerate(new_export_list):
            data.update({report: [new_filter_list[idx][0], new_filter_list[idx][1], new_filter_list[idx][2]]})

        with open("data/ExportSettings.json", "w") as f:
            json.dump(data, f, indent=4)

    def spec_component_settings(self):
        self.select_frame_by_name("frame_2")

        self.frame_label_2.grid_forget()
        self.browse_label_2.grid_forget()
        self.browse_entry_2.grid_forget()
        self.browse_button_2.grid_forget()
        self.settings_label_2_main.grid_forget()
        self.settings_button_2.grid_forget()
        self.close_checkbox_2.grid_forget()
        self.start_button_2.grid_forget()
        self.login_frame_2.grid_forget()
        self.login_label_2.grid_forget()
        self.user_name_2_label.grid_forget()
        self.user_name_2_entry.grid_forget()
        self.user_password_2_label.grid_forget()
        self.user_password_2_entry.grid_forget()

        self.vga_filter_list = []
        self.mb_filter_list = []
        self.psu_filter_list = []
        self.case_filter_list = []
        self.cooler_filter_list = []
        self.ssd_filter_list = []

        f = open("data/SPECComponentSettings.json",)
        f_data = json.load(f)
        for filter in f_data["VGA"]:
            self.vga_filter_list.append(filter)
        for filter in f_data["MB"]:
            self.mb_filter_list.append(filter)
        for filter in f_data["PSU"]:
            self.psu_filter_list.append(filter)
        for filter in f_data["CASE"]:
            self.case_filter_list.append(filter)
        for filter in f_data["COOLER"]:
            self.cooler_filter_list.append(filter)
        for filter in f_data["SSD"]:
            self.ssd_filter_list.append(filter)
        f.close()

        self.spec_frame_label_1 = customtkinter.CTkLabel(master=self.frame_2, text="SPEC komponentų nustatymai", font=("Calibri", 20))
        self.spec_frame_label_1.grid(row=0, column=0, columnspan=3, padx=20, pady=10)

        self.spec_filter_frame = customtkinter.CTkFrame(master=self.frame_2, corner_radius=0, fg_color="transparent")
        self.spec_filter_frame.grid(row=1, column=0, rowspan=5, columnspan=3, padx=20, pady=10)
        
        self.settings_label_2 = customtkinter.CTkLabel(master=self.spec_filter_frame, text="Komponentų filtrų nustatymai", font=("Calibri", 16))
        self.settings_label_2.grid(row=0, column=0, columnspan=3, padx=0, pady=(5, 2))

        self.vga_button = customtkinter.CTkButton(master=self.spec_filter_frame, text="VGA", width=75, font=("Calibri", 16), command=lambda: self.component_filter_settings(self.vga_button.cget("text")))
        self.vga_button.grid(row=1, column=0, columnspan=3, padx=(0, 0), pady=5)

        self.mb_button = customtkinter.CTkButton(master=self.spec_filter_frame, text="MB", width=75, font=("Calibri", 16), command=lambda: self.component_filter_settings(self.mb_button.cget("text")))
        self.mb_button.grid(row=2, column=0, columnspan=3, padx=(0, 0), pady=5)

        self.psu_button = customtkinter.CTkButton(master=self.spec_filter_frame, text="PSU", width=75, font=("Calibri", 16), command=lambda: self.component_filter_settings(self.psu_button.cget("text")))
        self.psu_button.grid(row=3, column=0, columnspan=3, padx=(0, 0), pady=5)

        self.case_button = customtkinter.CTkButton(master=self.spec_filter_frame, text="CASE", width=75, font=("Calibri", 16), command=lambda: self.component_filter_settings(self.case_button.cget("text")))
        self.case_button.grid(row=4, column=0, columnspan=3, padx=(0, 0), pady=5)

        self.cooler_button = customtkinter.CTkButton(master=self.spec_filter_frame, text="COOLER", width=75, font=("Calibri", 16), command=lambda: self.component_filter_settings(self.cooler_button.cget("text")))
        self.cooler_button.grid(row=5, column=0, columnspan=3, padx=(0, 0), pady=5)

        self.ssd_button = customtkinter.CTkButton(master=self.spec_filter_frame, text="SSD", width=75, font=("Calibri", 16), command=lambda: self.component_filter_settings(self.ssd_button.cget("text")))
        self.ssd_button.grid(row=6, column=0, columnspan=3, padx=(0, 0), pady=5)

        self.back_button_2_main = customtkinter.CTkButton(master=self.frame_2, text="<< Grįžti", width=100, font=("Calibri", 16), fg_color="gainsboro", text_color="black", hover_color="darkgrey", command=self.go_back_spec_component_settings)
        self.back_button_2_main.grid(row=7, column=0, columnspan=3, pady=10)
        
    
    def component_filter_settings(self, component):
        self.back_button_2_main.grid_forget()
        global spec_component_validation
        global spec_component_old_validation_error
        global spec_component_clicked_before
        global old_anchor_frame, old_new_component_label,old_new_component_entry, old_new_component_button, old_error_label_2, old_save_button_2, old_back_button_2
        if spec_component_validation:
            spec_component_old_validation_error.grid_forget()
            spec_component_validation = False
        
        if spec_component_clicked_before:
            for i in old_anchor_frame.winfo_children():
                i.destroy()
            old_anchor_frame.grid_forget()
            old_new_component_label.destroy()
            old_new_component_entry.destroy()
            old_new_component_button.destroy()
            old_error_label_2.destroy()
            old_save_button_2.destroy()
            old_back_button_2.destroy()
            spec_component_clicked_before = False

        self.component_filter_list = []
        if component == "VGA":
            self.component_filter_list = self.vga_filter_list
        elif component == "MB":
            self.component_filter_list = self.mb_filter_list
        elif component == "PSU":
            self.component_filter_list = self.psu_filter_list
        elif component == "CASE":
            self.component_filter_list = self.case_filter_list
        elif component == "COOLER":
            self.component_filter_list = self.cooler_filter_list
        elif component == "SSD":
            self.component_filter_list = self.ssd_filter_list

        self.vga_button.grid_configure(columnspan = 1)
        self.mb_button.grid_configure(columnspan = 1)
        self.psu_button.grid_configure(columnspan = 1)
        self.case_button.grid_configure(columnspan = 1)
        self.cooler_button.grid_configure(columnspan = 1)
        self.ssd_button.grid_configure(columnspan = 1)

        self.settings_label_2.configure(text=component + " nustatymai")
        component_entry_text = tk.StringVar(value="")
        self.anchor_frame = customtkinter.CTkScrollableFrame(master=self.spec_filter_frame, width=150, height=200)
        self.new_component_label = customtkinter.CTkLabel(master=self.spec_filter_frame, text="Pridėti nauja:", font=("Calibri", 16)) 
        self.new_component_entry = customtkinter.CTkEntry(master=self.spec_filter_frame, width=125, font=("Calibri", 16), textvariable=component_entry_text, placeholder_text="Filtras", border_color="black")
        self.new_component_button = customtkinter.CTkButton(master=self.spec_filter_frame, text="Pridėti", width=100, font=("Calibri", 16), command=lambda: self.add_component_filter(component_entry_text))

        self.anchor_frame.grid(row=1, column=1, rowspan=5, columnspan=1, padx=20, pady=10)
        self.new_component_label.grid(row=1, column=2, sticky="s", pady=(0,0))
        self.new_component_entry.grid(row=2, column=2, sticky="s", pady=(0,5))
        self.new_component_button.grid(row=3, column=2, sticky="n", pady=5)

        self.error_label_2 = customtkinter.CTkLabel(master=self.spec_filter_frame, text="Pavadinimas negali\nbūti trumpesnis\nnei 3 simboliai,\nir ne ilgesnis\nnei 15 simbolių", font=("Calibri", 14), text_color="red")
        self.error_label_2.grid(row=4, column=2, rowspan=2, sticky="s", pady=0)
        self.error_label_2.grid_forget()

        self.component_filter = []
        self.component_filter_delete = []
        for row_idx, filter_name in enumerate(self.component_filter_list):
            self.component_filter.append(customtkinter.CTkLabel(master=self.anchor_frame, text=filter_name, font=("Calibri", 14), width=85))
            self.component_filter_delete.append(customtkinter.CTkButton(master=self.anchor_frame, text="Ištrinti", width=30, font=("Calibri", 14), fg_color="darkgrey", text_color="black", hover_color="firebrick",
                                                            command=lambda row_idx=row_idx: self.delete_component_filter(self.component_filter[row_idx], self.component_filter_delete[row_idx], row_idx)))
            if filter_name != "TO_BE_DELETED":
                self.component_filter[row_idx].grid(row=row_idx, column=0, padx=(0, 0), pady=2)
                self.component_filter_delete[row_idx].grid(row=row_idx, column=1, padx=(10, 0), pady=2, sticky="ne")
        
        self.save_button_2 = customtkinter.CTkButton(master=self.frame_2, text="Išsaugoti", width=100, font=("Calibri", 16), command=self.save_component_filter_settings)
        self.save_button_2.grid(row=6, column=0, padx=(125, 5), pady=10, sticky="ne")
        self.back_button_2 = customtkinter.CTkButton(master=self.frame_2, text="<< Grįžti", width=100, font=("Calibri", 16), fg_color="gainsboro", text_color="black", hover_color="darkgrey", command=self.go_back_component_filter_settings)
        self.back_button_2.grid(row=6, column=1, padx=(5, 125), pady=10, sticky="nw")

        spec_component_clicked_before = True
        old_anchor_frame = self.anchor_frame
        old_new_component_label = self.new_component_label
        old_new_component_entry = self.new_component_entry
        old_new_component_button = self.new_component_button
        old_error_label_2 = self.error_label_2
        old_save_button_2 = self.save_button_2
        old_back_button_2 = self.back_button_2

    def add_component_filter(self, component_entry_text):
        component_filter_name = self.new_component_entry.get()
        if len(component_filter_name) > 2 and len(component_filter_name) < 15:
            self.component_filter_list.append(component_filter_name)
            row_idx = len(self.component_filter)
            self.new_component_entry.configure(border_color="black")
            component_entry_text.set("")
            self.error_label_2.grid_forget()

            self.component_filter.append(customtkinter.CTkLabel(master=self.anchor_frame, text=component_filter_name, font=("Calibri", 14), width=85))
            self.component_filter[row_idx].grid(row=row_idx, column=0, padx=(0, 0), pady=2)
            self.component_filter_delete.append(customtkinter.CTkButton(master=self.anchor_frame, text="Ištrinti", width=30, font=("Calibri", 14), fg_color="darkgrey", text_color="black", hover_color="firebrick",
                                                            command=lambda row_idx=row_idx: self.delete_component_filter(self.component_filter[row_idx], self.component_filter_delete[row_idx], row_idx)))
            self.component_filter_delete[row_idx].grid(row=row_idx, column=1, padx=(10, 0), pady=2, sticky="ne")
        else:
            global spec_component_validation
            global spec_component_old_validation_error
            spec_component_validation = True
            spec_component_old_validation_error = self.error_label_2
            self.new_component_entry.configure(border_color="red")
            self.error_label_2.grid(row=4, column=2, rowspan=2, sticky="s", pady=0)

    def delete_component_filter(self, specific_component_filter, specific_component_filter_delete, row_index):
        specific_component_filter.destroy()
        specific_component_filter_delete.destroy()
        self.component_filter_list[row_index] = "TO_BE_DELETED"
    
    def save_component_filter_settings(self):
        for i in self.anchor_frame.winfo_children():
            i.destroy()
        self.anchor_frame.grid_forget()
        self.new_component_label.destroy()
        self.new_component_entry.destroy()
        self.new_component_button.destroy()
        self.error_label_2.destroy()
        self.save_button_2.destroy()
        self.back_button_2.destroy()
        self.settings_label_2.configure(text="Komponentų filtrų nustatymai")
        self.vga_button.grid_configure(columnspan = 3)
        self.mb_button.grid_configure(columnspan = 3)
        self.psu_button.grid_configure(columnspan = 3)
        self.case_button.grid_configure(columnspan = 3)
        self.cooler_button.grid_configure(columnspan = 3)
        self.ssd_button.grid_configure(columnspan = 3)
        self.back_button_2_main.grid(row=6, column=0, columnspan=3, pady=10)


        self.vga_filter_list = [value for value in self.vga_filter_list if value != "TO_BE_DELETED"]
        self.mb_filter_list = [value for value in self.mb_filter_list if value != "TO_BE_DELETED"]
        self.psu_filter_list = [value for value in self.psu_filter_list if value != "TO_BE_DELETED"]
        self.case_filter_list = [value for value in self.case_filter_list if value != "TO_BE_DELETED"]
        self.cooler_filter_list = [value for value in self.cooler_filter_list if value != "TO_BE_DELETED"]
        self.ssd_filter_list = [value for value in self.ssd_filter_list if value != "TO_BE_DELETED"]

        data = {
            "VGA": self.vga_filter_list,
            "MB": self.mb_filter_list,
            "PSU": self.psu_filter_list,
            "CASE": self.case_filter_list,
            "COOLER": self.cooler_filter_list,
            "SSD": self.ssd_filter_list
        }
        with open("data/SPECComponentSettings.json", "w") as f:
            json.dump(data, f, indent=4)
    
    def go_back_component_filter_settings(self):
        self.anchor_frame.grid_forget()
        self.new_component_label.grid_forget()
        self.new_component_entry.grid_forget()
        self.new_component_button.grid_forget()
        self.error_label_2.grid_forget()
        self.save_button_2.grid_forget()
        self.back_button_2.grid_forget()
        self.back_button_2_main.grid_forget()
        self.spec_frame_label_1.grid_forget()
        self.spec_filter_frame.grid_forget()
        self.settings_label_2.grid_forget()
        self.vga_button.grid_forget()
        self.mb_button.grid_forget()
        self.psu_button.grid_forget()
        self.case_button.grid_forget()
        self.cooler_button.grid_forget()
        self.ssd_button.grid_forget()
        
        self.frame_label_2.grid(row=0, column=0, columnspan=3, padx=20, pady=10)
        self.browse_label_2.grid(row=1, column=0, columnspan=3, padx=0, pady=2)
        self.browse_entry_2.grid(row=2, column=0, padx=(40, 0), pady=5, sticky="e")
        self.browse_button_2.grid(row=2, column=1, padx=(5, 0), pady=5, sticky="w")
        self.settings_label_2_main.grid(row=3, column=0, columnspan=3, padx=0, pady=(10, 2))
        self.settings_button_2.grid(row=4, column=0, columnspan=3, padx=(0, 0), pady=5)
        self.close_checkbox_2.grid(row=5, column=0, columnspan=2, padx=(0, 0), pady=(15, 0))
        self.start_button_2.grid(row=10, column=0, columnspan=2, padx=(0, 0), pady=15)
        self.login_frame_2.grid(row=6, column=0, rowspan=3, columnspan=2, padx=20, pady=10)
        self.login_label_2.grid(row=0, column=0, columnspan=2, padx=0, pady=5)
        self.user_name_2_label.grid(row=1, column=0, padx=0, pady=3)
        self.user_name_2_entry.grid(row=1, column=1, padx=0, pady=3)
        self.user_password_2_label.grid(row=2, column=0, padx=0, pady=3)
        self.user_password_2_entry.grid(row=2, column=1, padx=0, pady=3)


    def go_back_spec_component_settings(self):
        self.back_button_2_main.grid_forget()
        self.spec_frame_label_1.grid_forget()
        self.spec_filter_frame.grid_forget()
        self.settings_label_2.grid_forget()
        self.vga_button.grid_forget()
        self.mb_button.grid_forget()
        self.psu_button.grid_forget()
        self.case_button.grid_forget()
        self.cooler_button.grid_forget()
        self.ssd_button.grid_forget()
        
        self.frame_label_2.grid(row=0, column=0, columnspan=3, padx=20, pady=10)
        self.browse_label_2.grid(row=1, column=0, columnspan=3, padx=0, pady=2)
        self.browse_entry_2.grid(row=2, column=0, padx=(40, 0), pady=5, sticky="e")
        self.browse_button_2.grid(row=2, column=1, padx=(5, 0), pady=5, sticky="w")
        self.settings_label_2_main.grid(row=3, column=0, columnspan=3, padx=0, pady=(10, 2))
        self.settings_button_2.grid(row=4, column=0, columnspan=3, padx=(0, 0), pady=5)
        self.close_checkbox_2.grid(row=5, column=0, columnspan=2, padx=(0, 0), pady=(15, 0))
        self.start_button_2.grid(row=10, column=0, columnspan=2, padx=(0, 0), pady=15)
        self.login_frame_2.grid(row=6, column=0, rowspan=3, columnspan=2, padx=20, pady=10)
        self.login_label_2.grid(row=0, column=0, columnspan=2, padx=0, pady=5)
        self.user_name_2_label.grid(row=1, column=0, padx=0, pady=3)
        self.user_name_2_entry.grid(row=1, column=1, padx=0, pady=3)
        self.user_password_2_label.grid(row=2, column=0, padx=0, pady=3)
        self.user_password_2_entry.grid(row=2, column=1, padx=0, pady=3)

    def initiate_weekly_report_module(self):
        self.geometry("700x550+1000+0")
        self.open_console_frame_1()
        wrp.main(self, self.user_name_1_entry.get(), self.user_password_1_entry.get())
        if self.spec_checked.get() == 1:
            wmrp.main(self, self.user_name_2_entry.get(), self.user_password_2_entry.get())    
        
        if self.close_checked_1.get() == 1:
            main_processName = "EFrame2"
            main_pid = application.process_from_module(module = main_processName)
            main_process = application.Application().connect(process = main_pid, timeout=10)
            main_process.Programa.close()

    def initiate_spec_weekly_report_module(self):
        self.geometry("700x550+1000+0")
        self.open_console_frame_2()
        wmrp.main(self, self.user_name_2_entry.get(), self.user_password_2_entry.get())
        
        if self.close_checked_2.get() == 1:
            main_processName = "EFrame2"
            main_pid = application.process_from_module(module = main_processName)
            main_process = application.Application().connect(process = main_pid, timeout=10)
            main_process.Programa.close()

    def initiate_spec_monthly_report_module(self):
        self.geometry("700x550+1000+0")
        self.open_console_frame_3()
        mmrp.main(self, self.user_name_3_entry.get(), self.user_password_3_entry.get())
        
        if self.close_checked_3.get() == 1:
            main_processName = "EFrame2"
            main_pid = application.process_from_module(module = main_processName)
            main_process = application.Application().connect(process = main_pid, timeout=10)
            main_process.Programa.close()

    def open_console_frame_1(self):
        self.frame_label_1.grid_forget()
        self.browse_label_1.grid_forget()
        self.browse_entry_1.grid_forget()
        self.browse_button_1.grid_forget()
        self.settings_label_1.grid_forget()
        self.manufacturer_button_1.grid_forget()
        self.settings_button_1.grid_forget()
        self.spec_checkbox_1.grid_forget()
        self.close_checkbox_1.grid_forget()
        self.start_button_1.grid_forget()
        self.login_frame_1.grid_forget()
        self.login_label_1.grid_forget()
        self.user_name_1_label.grid_forget()
        self.user_name_1_entry.grid_forget()
        self.user_password_1_label.grid_forget()
        self.user_password_1_entry.grid_forget()

        self.console_frame = customtkinter.CTkFrame(master=self.frame_1, width=400,  height=450)
        self.console_frame.grid(row=0, column=0, columnspan=2, padx=20, pady=10)
        self.console_textbox = customtkinter.CTkTextbox(master=self.console_frame, font=("Calibri", 14),width=400, height=450)
        self.console_textbox.grid(row=0, column=0, padx=0, pady=0)
        self.console_back_button = customtkinter.CTkButton(master=self.frame_1, text="<< Grįžti", width=100, font=("Calibri", 16), fg_color="gainsboro", text_color="black", hover_color="darkgrey", command=self.go_back_console_1)
        self.console_back_button.grid(row=1, column=0, padx=(100,20), pady=5, sticky="ne")
        self.console_stop_button = customtkinter.CTkButton(master=self.frame_1, text="STOP", width=100, font=("Calibri", 16), fg_color="firebrick", text_color="white", hover_color="red", command=self.stop_console)
        self.console_stop_button.grid(row=1, column=1, padx=(20,100), pady=5, sticky="nw")
    
    def go_back_console_1(self):
        self.console_textbox.delete("1.0", "end")

        self.console_frame.grid_forget()
        self.console_textbox.grid_forget()
        self.console_back_button.grid_forget()
        self.console_stop_button.grid_forget()

        self.frame_label_1.grid(row=0, column=0, columnspan=2, padx=20, pady=10)
        self.browse_label_1.grid(row=1, column=0, columnspan=2, padx=0, pady=2)
        self.browse_entry_1.grid(row=2, column=0, padx=(40, 0), pady=5, sticky="e")
        self.browse_button_1.grid(row=2, column=1, padx=(5, 0), pady=5, sticky="w")
        self.settings_label_1.grid(row=3, column=0, columnspan=2, padx=0, pady=(10, 2))
        self.manufacturer_button_1.grid(row=4, column=0, columnspan=2, padx=(0, 0), pady=5)
        self.settings_button_1.grid(row=5, column=0, columnspan=2, padx=(0, 0), pady=5)
        self.spec_checkbox_1.grid(row=6, column=0, columnspan=2, padx=(0, 0), pady=(20, 5))
        self.close_checkbox_1.grid(row=7, column=0, columnspan=2, padx=(0, 0), pady=5)
        self.start_button_1.grid(row=11, column=0, columnspan=2, padx=(0, 0), pady=15)
        self.login_frame_1.grid(row=8, column=0, rowspan=3, columnspan=2, padx=20, pady=10)
        self.login_label_1.grid(row=0, column=0, columnspan=2, padx=0, pady=5)
        self.user_name_1_label.grid(row=1, column=0, padx=0, pady=3)
        self.user_name_1_entry.grid(row=1, column=1, padx=0, pady=3)
        self.user_password_1_label.grid(row=2, column=0, padx=0, pady=3)
        self.user_password_1_entry.grid(row=2, column=1, padx=0, pady=3)

    def open_console_frame_2(self):
        self.frame_label_2.grid_forget()
        self.browse_label_2.grid_forget()
        self.browse_entry_2.grid_forget()
        self.browse_button_2.grid_forget()
        self.settings_label_2_main.grid_forget()
        self.settings_button_2.grid_forget()
        self.close_checkbox_2.grid_forget()
        self.start_button_2.grid_forget()
        self.login_frame_2.grid_forget()
        self.login_label_2.grid_forget()
        self.user_name_2_label.grid_forget()
        self.user_name_2_entry.grid_forget()
        self.user_password_2_label.grid_forget()
        self.user_password_2_entry.grid_forget()

        self.console_frame = customtkinter.CTkFrame(master=self.frame_2, width=400,  height=450)
        self.console_frame.grid(row=0, column=0, columnspan=2, padx=20, pady=10)
        self.console_textbox = customtkinter.CTkTextbox(master=self.console_frame, font=("Calibri", 14),width=400, height=450)
        self.console_textbox.grid(row=0, column=0, padx=0, pady=0)
        self.console_back_button = customtkinter.CTkButton(master=self.frame_2, text="<< Grįžti", width=100, font=("Calibri", 16), fg_color="gainsboro", text_color="black", hover_color="darkgrey", command=self.go_back_console_2)
        self.console_back_button.grid(row=1, column=0, padx=(100,20), pady=5, sticky="ne")
        self.console_stop_button = customtkinter.CTkButton(master=self.frame_2, text="STOP", width=100, font=("Calibri", 16), fg_color="firebrick", text_color="white", hover_color="red", command=self.stop_console)
        self.console_stop_button.grid(row=1, column=1, padx=(20,100), pady=5, sticky="nw")

    def go_back_console_2(self):
        self.console_textbox.delete("1.0", "end")
        
        self.console_frame.grid_forget()
        self.console_textbox.grid_forget()
        self.console_back_button.grid_forget()
        self.console_stop_button.grid_forget()

        self.frame_label_2.grid(row=0, column=0, columnspan=3, padx=20, pady=10)
        self.browse_label_2.grid(row=1, column=0, columnspan=3, padx=0, pady=2)
        self.browse_entry_2.grid(row=2, column=0, padx=(40, 0), pady=5, sticky="e")
        self.browse_button_2.grid(row=2, column=1, padx=(5, 0), pady=5, sticky="w")
        self.settings_label_2_main.grid(row=3, column=0, columnspan=3, padx=0, pady=(10, 2))
        self.settings_button_2.grid(row=4, column=0, columnspan=3, padx=(0, 0), pady=5)
        self.close_checkbox_2.grid(row=5, column=0, columnspan=2, padx=(0, 0), pady=(15, 0))
        self.start_button_2.grid(row=10, column=0, columnspan=2, padx=(0, 0), pady=15)
        self.login_frame_2.grid(row=6, column=0, rowspan=3, columnspan=2, padx=20, pady=10)
        self.login_label_2.grid(row=0, column=0, columnspan=2, padx=0, pady=5)
        self.user_name_2_label.grid(row=1, column=0, padx=0, pady=3)
        self.user_name_2_entry.grid(row=1, column=1, padx=0, pady=3)
        self.user_password_2_label.grid(row=2, column=0, padx=0, pady=3)
        self.user_password_2_entry.grid(row=2, column=1, padx=0, pady=3)


    def open_console_frame_3(self):
        self.frame_label_3.grid_forget()
        self.browse_label_3.grid_forget()
        self.browse_entry_3.grid_forget()
        self.browse_button_3.grid_forget()
        self.settings_label_3_main.grid_forget()
        self.settings_button_3.grid_forget()
        self.close_checkbox_3.grid_forget()
        self.start_button_3.grid_forget()
        self.login_frame_3.grid_forget()
        self.login_label_3.grid_forget()
        self.user_name_3_label.grid_forget()
        self.user_name_3_entry.grid_forget()
        self.user_password_3_label.grid_forget()
        self.user_password_3_entry.grid_forget()

        self.console_frame = customtkinter.CTkFrame(master=self.frame_3, width=400,  height=450)
        self.console_frame.grid(row=0, column=0, columnspan=2, padx=20, pady=10)
        self.console_textbox = customtkinter.CTkTextbox(master=self.console_frame, font=("Calibri", 14),width=400, height=450)
        self.console_textbox.grid(row=0, column=0, padx=0, pady=0)
        self.console_back_button = customtkinter.CTkButton(master=self.frame_3, text="<< Grįžti", width=100, font=("Calibri", 16), fg_color="gainsboro", text_color="black", hover_color="darkgrey", command=self.go_back_console_3)
        self.console_back_button.grid(row=1, column=0, padx=(100,20), pady=5, sticky="ne")
        self.console_stop_button = customtkinter.CTkButton(master=self.frame_3, text="STOP", width=100, font=("Calibri", 16), fg_color="firebrick", text_color="white", hover_color="red", command=self.stop_console)
        self.console_stop_button.grid(row=1, column=1, padx=(20,100), pady=5, sticky="nw")

    def go_back_console_3(self):
        self.console_textbox.delete("1.0", "end")
        
        self.console_frame.grid_forget()
        self.console_textbox.grid_forget()
        self.console_back_button.grid_forget()
        self.console_stop_button.grid_forget()

        self.frame_label_3.grid(row=0, column=0, columnspan=3, padx=20, pady=10)
        self.browse_label_3.grid(row=1, column=0, columnspan=3, padx=0, pady=2)
        self.browse_entry_3.grid(row=2, column=0, padx=(40, 0), pady=5, sticky="e")
        self.browse_button_3.grid(row=2, column=1, padx=(5, 0), pady=5, sticky="w")
        self.settings_label_3_main.grid(row=3, column=0, columnspan=3, padx=0, pady=(10, 2))
        self.settings_button_3.grid(row=4, column=0, columnspan=3, padx=(0, 0), pady=5)
        self.close_checkbox_3.grid(row=5, column=0, columnspan=2, padx=(0, 0), pady=(15, 0))
        self.start_button_3.grid(row=10, column=0, columnspan=2, padx=(0, 0), pady=15)
        self.login_frame_3.grid(row=6, column=0, rowspan=3, columnspan=2, padx=20, pady=10)
        self.login_label_3.grid(row=0, column=0, columnspan=2, padx=0, pady=5)
        self.user_name_3_label.grid(row=1, column=0, padx=0, pady=3)
        self.user_name_3_entry.grid(row=1, column=1, padx=0, pady=3)
        self.user_password_3_label.grid(row=2, column=0, padx=0, pady=3)
        self.user_password_3_entry.grid(row=2, column=1, padx=0, pady=3)


    def stop_console(self):
        os._exit(0)

    def write_to_gui(self, text):
        now = datetime.datetime.now()
        current_time = now.strftime("%H:%M:%S")
        self.console_textbox.insert("end", "[" + current_time + "]: " + text + '\n')
        self.console_textbox.see("end")

if __name__ == "__main__":
    app = App()
    app.mainloop()
    