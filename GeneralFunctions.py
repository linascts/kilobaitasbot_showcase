import UserControl
import pyautogui, time, win32api, win32clipboard
from openpyxl import Workbook
from pywinauto import application
from pywinauto.keyboard import send_keys

global is_lt_lang
is_lt_lang = False

def print_ui(text, ui_self):
    print(text)
    UserControl.App.write_to_gui(ui_self, text)

def find_current_dpi(ui_self, image, conf):
    if pyautogui.locateOnScreen("images/100/" + image, confidence=conf):
        return "images/100/"
    elif pyautogui.locateOnScreen("images/125/" + image, confidence=conf):
        print_ui("Nustatyta, kad pagrindinis ekranas naudoja 25 procentų padidintus langus ant originalios rezoliucijos.", ui_self)
        return "images/125/"
    elif pyautogui.locateOnScreen("images/150/" + image, confidence=conf):
        print_ui("Nustatyta, kad pagrindinis ekranas naudoja 50 procentų padidintus langus ant originalios rezoliucijos.", ui_self)
        return "images/150/"
    elif pyautogui.locateOnScreen("images/175/" + image, confidence=conf):
        print_ui("Nustatyta, kad pagrindinis ekranas naudoja 75 procentų padidintus langus ant originalios rezoliucijos.", ui_self)
        return "images/175/"
    else:
        print_ui("Nerastas meniu, stabdoma programa. Patikrinkite rezoliucijos (scaling) procenta, jis gali būti nuo 100 iki 175.", ui_self)
        quit()

def connect_sequence(ui_self, main_process, user_name, user_password):
    time.sleep(1)
    login_dlg = main_process.window(class_name="TFormPasswordDialog")
    login_dlg.set_focus()
    textbox_username = login_dlg.Edit1
    textbox_password = login_dlg.Edit2
    button_confirm = login_dlg.Button2
    textbox_username.type_keys(user_name, with_spaces = True)
    textbox_password.type_keys(user_password, with_spaces = True)
    button_confirm.click_input()
    send_keys("{ENTER}")
    time.sleep(3)

    # Checks if the information input is valid
    if login_dlg.exists:
        try:
            login_dlg.close()
            time.sleep(1)
            main_process.Programa.close()
            print_ui("Įvesti neteisingi duomenys, prašome pabandyti dar karta", ui_self)
            quit()
        except Exception:
            print_ui("Pilnai palaukiam kol užsikraus įmonės programos SQL...", ui_self)
            main_process.Programa.wait("exists enabled visible ready", timeout=120)
            time.sleep(2)
            pass

def monitor_set_position(ui_self, main_dlg):
    monitors = win32api.EnumDisplayMonitors()
    total_monitors = len(monitors)
    if total_monitors > 1:
        print_ui("Programa aptiko kad naudojate daugiau nei viena monitorių, automatizavimo procesas vyksta tik ant vieno monitoriaus...", ui_self)
        i = 0
        while i < total_monitors:
            print(str(win32api.GetMonitorInfo(monitors[i][0])), ui_self)
            i += 1
        time.sleep(1)
        print_ui("Įmonės programos langas perkeliamas į pagrindinį ekrana...", ui_self)
        main_dlg.move_window(x=0, y=0, width=1000, height=800, repaint=True)
        time.sleep(1)
    else:
        main_dlg.move_window(x=0, y=0, width=1000, height=800, repaint=True)
        time.sleep(1)

def locate_and_click(image, conf_value):
    try:
        x, y = pyautogui.locateCenterOnScreen(image, confidence=conf_value)
        pyautogui.moveTo(x, y)
        pyautogui.leftClick()
    except Exception:
        lt_image = image.replace(".png", "-lt.png")
        x, y = pyautogui.locateCenterOnScreen(lt_image, confidence=conf_value)
        pyautogui.moveTo(x, y)
        pyautogui.leftClick()
        global is_lt_lang
        is_lt_lang = True

def input_date_by_locale(textbox_start, textbox_end, translate_date_begin, translate_date_end, translate_date_begin_lt, translate_date_end_lt):
    if is_lt_lang:
        textbox_start.type_keys(translate_date_begin_lt)
        time.sleep(1)
        textbox_end.type_keys(translate_date_end_lt)
    else:
        textbox_start.type_keys(translate_date_begin)
        time.sleep(1)
        textbox_end.type_keys(translate_date_end)

def navigate_menu(ui_self, first_click, first_click_conf, second_click, second_click_conf, third_click, third_click_conf, main_dlg):
    lt_second_click = second_click.replace(".png", "-lt.png")
    lt_third_click = third_click.replace(".png", "-lt.png")
    locate_and_click(first_click, first_click_conf)
    check_dropdown = pyautogui.locateOnScreen(second_click, confidence=second_click_conf)
    if check_dropdown == None:
        check_dropdown = pyautogui.locateOnScreen(lt_second_click, confidence=second_click_conf)

    if check_dropdown == None:
        print_ui("Programa nerado pasirinkimo...", ui_self)
        time.sleep(1)
        print_ui("Palaukiam ir bandom iš naujo...", ui_self)
        time.sleep(3)
        check_dropdown = pyautogui.locateOnScreen(second_click, confidence=second_click_conf)
        if check_dropdown == None:
            check_dropdown = pyautogui.locateOnScreen(lt_second_click, confidence=second_click_conf)
        
        if check_dropdown == None:
            print_ui("Programa nerado pasirinkimo...", ui_self)
            time.sleep(1)
            print_ui("Uždarom programa.", ui_self)
            main_dlg.close()
            quit()
        else:
            locate_and_click(second_click, second_click_conf)
    else:
        locate_and_click(second_click, second_click_conf)
    
    check_second_dropdown = pyautogui.locateOnScreen(third_click, confidence=third_click_conf)
    if check_second_dropdown == None:
        check_second_dropdown = pyautogui.locateOnScreen(lt_third_click, confidence=third_click_conf)
    
    if check_second_dropdown == None:
        print_ui("Programa nerado pasirinkimo...", ui_self)
        time.sleep(1)
        print_ui("Palaukiam ir bandom iš naujo...", ui_self)
        time.sleep(3)
        check_second_dropdown = pyautogui.locateOnScreen(third_click, confidence=third_click_conf)
        if check_second_dropdown == None:
            check_second_dropdown = pyautogui.locateOnScreen(lt_third_click, confidence=third_click_conf)

        if check_second_dropdown == None:
            print_ui("Programa nerado pasirinkimo...", ui_self)
            time.sleep(1)
            print_ui("Uždarom programa.", ui_self)
            main_dlg.close()
            quit()
        else:
            locate_and_click(third_click, third_click_conf)
    else:
        locate_and_click(third_click, third_click_conf)

def refresh_data(scale):
    time.sleep(1)
    send_keys("^{F5}")
    time.sleep(1)
    for _ in range(100):
        time.sleep(1)
        a = pyautogui.locateOnScreen(scale + "atnaujinami-duomenys.png", confidence=0.8)
        if a == None:
            time.sleep(1)
            b = pyautogui.locateOnScreen(scale + "row-arrow.png", confidence=0.8)
            if b != None:
                locate_and_click(scale + "row-arrow.png", 0.8)
                time.sleep(1)
                break

def datagrid_search(filter, filter_title, filter_title_lt, clear_filter, main_process):
    filter_column = 0
    if len(filter) > 1 or clear_filter is True:
        for column in range(100):
            send_keys("{F8}")
            time.sleep(1)
            main_process.top_window().set_focus()
            if main_process.top_window().window_text() == filter_title or main_process.top_window().window_text() == filter_title_lt:
                if main_process.top_window().window_text() == filter_title:
                    filter_dlg = main_process.window(title_re=filter_title)
                elif main_process.top_window().window_text() == filter_title_lt:
                    filter_dlg = main_process.window(title_re=filter_title_lt)
                textbox_filter = filter_dlg.Edit
                if clear_filter is False:
                    textbox_filter.type_keys(filter, with_spaces=True)
                elif clear_filter is True:
                    textbox_filter.type_keys('{BACKSPACE}')
                send_keys("{ESC}")
                filter_column = column
                break
            else:
                send_keys("{ESC}")
                send_keys("{RIGHT}")
            column += 1

    main_process.Programa.set_focus()
    column = 0
    while column <= filter_column:
        send_keys("{LEFT}")
        column += 1

def clipboard_to_excel(company, report_type, export_path):
    win32clipboard.OpenClipboard()
    clipboard_data = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    clipboard_data = clipboard_data.replace('\t ', '\t')
    clipboard_data = [i.split('\t') for i in clipboard_data.split('\r\n')]
    wb = Workbook()
    ws = wb.active
    for row, row_data in enumerate(clipboard_data, start=1):
        for col, cell_data in enumerate(row_data, start=1):
            ws.cell(row=row, column=col, value=cell_data)
    wb.save(export_path + "\\" + company + "_" + report_type + ".xlsx")


def close_window(ui_self, class_name, main_dlg):
    for _ in range(20):
        time.sleep(1)
        window = main_dlg.child_window(class_name=class_name, found_index=0, visible_only=True)
        if window.exists(timeout=1):
            main_dlg.child_window(class_name=class_name, found_index=0, visible_only=True).close()
            print_ui("Uždaromas langas su kodu " + str(class_name), ui_self)
        else:
            break

def connect(ui_self, user_name, user_password):
    main_processName = "EFrame2"
    try: 
        detect_viber(ui_self)
        main_pid = application.process_from_module(module = main_processName)
        main_process = application.Application().connect(process = main_pid, timeout=10)
        print_ui("Proceso ID: " + str(main_pid), ui_self)
        main_process.Programa.set_focus()
        main_dlg = main_process.window(class_name="TFormMDI")
        monitor_set_position(ui_self, main_dlg)
    except Exception:
        send_keys('^%e') # Klaviatūros kombinacija CTRL + ALT + E (ATIDARO įmonės programa)
        print_ui("Paleidžiama įmonės programa...", ui_self)
        time.sleep(5)
        main_pid = application.process_from_module(module = main_processName)
        print_ui("Proceso ID: " + str(main_pid), ui_self)
        main_process = application.Application().connect(process = main_pid, timeout=10)
        main_dlg = main_process.window(class_name="TFormMDI") 
        connect_sequence(ui_self, main_process, user_name, user_password)
        monitor_set_position(ui_self, main_dlg)
        pass

    return main_process, main_dlg

def detect_viber(ui_self):
    viber_processName = "Viber"
    try:
        viber_pid = application.process_from_module(module = viber_processName)
        application.Application().connect(process = viber_pid, timeout=3)
        print_ui("Aptikta Viber susirašinėjimo programa jūsų kompiuteryje kuri neleidžia startuoti Euroskaitai. Prašome išjungti Viber ir tada vykdyti automatizacijos procesa.", ui_self)
        quit()
    except Exception:
        pass