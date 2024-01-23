import GeneralFunctions as gf
import time, os, pyautogui, shutil, openpyxl, pywinauto, json
from datetime import datetime, timedelta
from pywinauto.keyboard import send_keys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from webdriver_auto_update import check_driver

# Functions

def create_directories(ui_self, main_path, main_directory_name, export_path, report_path, export_directory_name, report_directory_name):
    if os.path.exists(main_path):
        gf.print_ui("Praeitos savaitės aplankas jau buvo sukurtas, informacija talpinam į ta patį aplanka.", ui_self)
        pass
    else:
        os.mkdir(main_path)
        gf.print_ui("Nurodytam take, sukurtas naujas aplankas pavadinimu: " + main_directory_name, ui_self)
    
    if os.path.exists(export_path):
        shutil.rmtree(export_path)
    
    if os.path.exists(report_path):
        shutil.rmtree(report_path)

    os.mkdir(export_path)
    gf.print_ui("Aplanke [" + main_directory_name + "] sukurtas naujas aplankas: " + export_directory_name, ui_self)

    os.mkdir(report_path)
    gf.print_ui("Aplanke [" + main_directory_name + "] sukurtas naujas aplankas: " + report_directory_name, ui_self)

def check_if_by_group(ui_self, scale):
    for _ in range(15):
        send_keys("{UP}")

    check_dropdown = pyautogui.locateOnScreen(scale + "pagal-prekiu-grupes.png", confidence=0.8)
    if check_dropdown == None:
        gf.print_ui("Nerado pagal prekių grupes pasirinkimo...", ui_self)
        for _ in range(15):
            check_dropdown = pyautogui.locateOnScreen(scale + "pagal-prekiu-grupes.png", confidence=0.8)
            if check_dropdown == None:
                send_keys("{DOWN}")
            else:
                gf.locate_and_click(scale + "pagal-prekiu-grupes.png", 0.8)
                break
    else:
        gf.locate_and_click(scale + "pagal-prekiu-grupes.png", 0.8)

def stock_export(company, filter_by_group, filter_by_sku, filter_by_title, main_process, export_path, scale): 
    gf.datagrid_search(filter_by_group, "Grupe", "Grupė", False, main_process)
    gf.datagrid_search(filter_by_sku, "Kodas", "Kodas", False, main_process)
    gf.datagrid_search(filter_by_title, "Pavadinimas", "Pavadinimas", False, main_process)
    send_keys("{ESC}")
    gf.locate_and_click(scale + "export-button.png", 0.8)
    gf.datagrid_search("", "Grupe", "Grupė", True, main_process)
    gf.datagrid_search("", "Kodas", "Kodas", True, main_process)
    gf.datagrid_search("", "Pavadinimas", "Pavadinimas", True, main_process)
    gf.clipboard_to_excel(company.lower(), "stock", export_path)

def sales_export(company, filter_by_group, filter_by_sku, filter_by_title, main_process, export_path, scale):
    if company == "unknown_vga" or company == "unknown_nb":
        gf.datagrid_search(filter_by_group, "Grupe", "Grupė", False, main_process)
        gf.datagrid_search(filter_by_title, "Vertybes pavadinimas", "Vertybės pavadinimas", False, main_process)
        send_keys("{ESC}")
        gf.locate_and_click(scale + "export-button.png", 0.8)
        gf.datagrid_search("", "Grupe", "Grupė", True, main_process)
        gf.datagrid_search("", "Vertybes pavadinimas", "Vertybės pavadinimas", True, main_process)
    else:
        gf.datagrid_search(filter_by_group, "Grupe", "Grupė", False, main_process)
        gf.datagrid_search(filter_by_sku, "Pr.kodas", "Pr.kodas", False, main_process)
        gf.datagrid_search(filter_by_title, "Pr.pavad.", "Pr.pavad.", False, main_process)
        send_keys("{ESC}")
        gf.locate_and_click(scale + "export-button.png", 0.8)
        gf.datagrid_search("", "Grupe", "Grupė", True, main_process)
        gf.datagrid_search("", "Pr.kodas", "Pr.kodas", True, main_process)
        gf.datagrid_search("", "Pr.pavad.", "Pr.pavad.", True, main_process)
    
    gf.clipboard_to_excel(company.lower(), "sales", export_path)

def pc_sales_export(company, filter_by_sku, filter_by_title, main_process, export_path, scale):
    gf.datagrid_search(filter_by_sku, "ˇaliava", "Žaliava", False, main_process)
    gf.datagrid_search(filter_by_title, "ˇ.pavadinimas", "Ž.pavadinimas", False, main_process)
    send_keys("{ESC}")
    gf.locate_and_click(scale + "export-button.png", 0.8)
    gf.datagrid_search("", "ˇaliava", "Žaliava", True, main_process)
    gf.datagrid_search("", "ˇ.pavadinimas", "Ž.pavadinimas", True, main_process)
    gf.clipboard_to_excel(company.lower(), "zal", export_path)

def check_manufacturer(title_value, template_sheet_obj, c_row, manufacturer_column, manufacturer_alignment, manufacturers):
    for manufacturer in manufacturers:
        if manufacturer.lower() in title_value.lower():
            template_cell_obj = template_sheet_obj.cell(row = c_row, column = manufacturer_column)
            template_cell_obj.alignment = openpyxl.styles.Alignment(horizontal=manufacturer_alignment)
            template_cell_obj.value = manufacturer

def insert_column_values(from_export, export_sheet_obj, export_max_row, export_column, template_sheet_obj, template_column, template_alignment, template_value, value_to_float, with_manufacturer, manufacturer_column, manufacturer_alignment, value_to_datetime, select_row, manufacturers):
    if from_export is True:
        for i in range(2, export_max_row + 1):
            cell_obj = export_sheet_obj.cell(row = i, column = export_column)
            template_cell_obj = template_sheet_obj.cell(row = i + select_row, column = template_column)
            template_cell_obj.alignment = openpyxl.styles.Alignment(horizontal=template_alignment)
            if value_to_float is True:
                if isinstance(cell_obj.value, str):
                    temp_value = cell_obj.value
                    if '\xa0' in cell_obj.value:
                        temp_value = cell_obj.value.replace(u'\xa0', u'')
                        
                    if ',000000' in temp_value:
                        template_cell_obj.value = float(temp_value.replace(',000000', ''))
                    elif ',000' in temp_value:
                        template_cell_obj.value = float(temp_value.replace(',000', ''))
                    elif ',' in temp_value and '.' in temp_value:
                        template_cell_obj.value = float(temp_value.replace(',', ''))
                    elif ',' in temp_value: 
                        template_cell_obj.value = float(temp_value.replace(',', '.'))
                    else:
                        template_cell_obj.value = float(temp_value)


            elif value_to_datetime is True:
                if isinstance(cell_obj.value, str):
                    try:
                        date_object = datetime.strptime(cell_obj.value, "%m/%d/%Y")
                        template_cell_obj.value = str(date_object.strftime("%Y.%m.%d"))
                    except Exception:
                        template_cell_obj.value = cell_obj.value
                        pass
            else:
                template_cell_obj.value = cell_obj.value
                if with_manufacturer is True:
                    if isinstance(cell_obj.value, str):
                        check_manufacturer(cell_obj.value, template_sheet_obj, i + select_row, manufacturer_column, manufacturer_alignment, manufacturers)
    else:
        for i in range(2, export_max_row + 1):
            template_cell_obj = template_sheet_obj.cell(row = i + select_row, column = template_column)
            template_cell_obj.alignment = openpyxl.styles.Alignment(horizontal=template_alignment)
            template_cell_obj.value = template_value

def scrap_customer_info(export_sheet_obj, export_max_row, export_column, template_sheet_obj, template_column, template_alignment, template_value, driver, select_row):
    time.sleep(2)

    for i in range(2, export_max_row + 1):
        cell_obj = export_sheet_obj.cell(row = i, column = export_column)
        template_cell_obj = template_sheet_obj.cell(row = i + select_row, column = template_column)
        template_cell_obj.alignment = openpyxl.styles.Alignment(horizontal=template_alignment)
        title_obj = template_sheet_obj.cell(row = i + select_row, column = template_column - 1)
        title_obj.alignment = openpyxl.styles.Alignment(horizontal="left")
        if cell_obj.value == None:
            template_cell_obj.value = template_value
        else:
            try:
                time.sleep(4)
                sas_close = driver.find_element("xpath", "//div[contains(@id, 'sas_closeButton_')]")
                sas_close.click()
            except Exception:
                pass
    
            try:
                cookie_close = driver.find_element("xpath", "//div[@id='cookiescript_close']")
                cookie_close.click()
            except Exception:
                pass
            
            try:
                input_code = driver.find_element("xpath", "//input[@id='code']")
                input_code.clear()
                input_code.send_keys(str(cell_obj.value))
                time.sleep(1)
                input_code.send_keys(Keys.RETURN)
                time.sleep(2)
            except Exception:
                pass

            try:
                company_info = driver.find_element("xpath","//div[@class='address']")
                company_info = str(company_info.get_attribute("innerHTML"))
                if "LT-" in company_info:
                    start_of_code = company_info.index("LT-")
                    postal_code = company_info[start_of_code:start_of_code + 8]
                    template_cell_obj.value = postal_code
                    company_title = driver.find_element("xpath", "//a[@class='company-title d-block']")
                    title_obj.value = str(company_title.get_attribute("title"))
                else:
                    template_cell_obj.value = template_value
            except Exception:
                template_cell_obj.value = template_value
                pass

def create_report(company, report_type, manufacturers, export_path, today_date, prior_week_start, prior_week_end, report_path):
    path = export_path + "\\" + company.lower() + "_" +  report_type.lower() + ".xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet_obj = wb_obj.active
    max_col = sheet_obj.max_column
    m_row = sheet_obj.max_row

    for i in range(1, max_col + 1):
        cell_obj = sheet_obj.cell(row = 1, column = i)
        if cell_obj.value == "Pr.kodas" or cell_obj.value == "Kodas":
            column_of_sku = i
        elif cell_obj.value == "Pr.pavad." or cell_obj.value == "Pavadinimas":
            column_of_title = i
        elif cell_obj.value == "Sand.likutis" or cell_obj.value == "Kiekis":
            column_of_quantity = i
           
    file_date = None
    excel_date = None
    template_wb_obj = None

    if "stock" in report_type:
        template_wb_obj = openpyxl.load_workbook("templates/stock_example.xlsx")
        file_date = today_date[5] + today_date[6] + today_date[8] + today_date[9]
        excel_date = "Inventory " + today_date[5] + today_date[6] + "." + today_date[8] + today_date[9]
    elif "sales" in report_type:
        template_wb_obj = openpyxl.load_workbook("templates/sales_example.xlsx")
        file_date = prior_week_start[5] + prior_week_start[6] + prior_week_start[8] + prior_week_start[9] + "-" + prior_week_end[5] + prior_week_end[6] + prior_week_end[8] + prior_week_end[9]
        excel_date = "Sales qty " + prior_week_start[5] + prior_week_start[6] + "." + prior_week_start[8] + prior_week_start[9] + "-" + prior_week_end[5] + prior_week_end[6] + "." + prior_week_end[8] + prior_week_end[9]
    
    template_sheet_obj = template_wb_obj.active
    template_sheet_obj['A1'] = company.upper() + " P/N"
    template_sheet_obj['C1'] = excel_date

    insert_column_values(True, sheet_obj, m_row, column_of_sku, template_sheet_obj, 1, "left", None, False, False, None, None, False, 0, manufacturers)
    insert_column_values(True, sheet_obj, m_row, column_of_title, template_sheet_obj, 2, "left", None, False, False, None, None, False, 0, manufacturers)
    insert_column_values(True, sheet_obj, m_row, column_of_quantity, template_sheet_obj, 3, "center", None, True, False, None, None, False, 0, manufacturers)

    if "sales" in report_type:
        maximum_rows_in_report = template_sheet_obj.max_row
        template_sheet_obj.delete_rows(maximum_rows_in_report - 1)
        
    template_wb_obj.save(report_path + "\\" + company.lower() + "_" + report_type.lower() + file_date + ".xlsx")

def create_unknown_stock_report(manufacturers, export_path, today_date, report_path, stock_row):
    path = export_path + "\\" + "unknown_stock.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet_obj = wb_obj.active
    max_col = sheet_obj.max_column
    m_row = sheet_obj.max_row 
    for i in range(1, max_col + 1):
        cell_obj = sheet_obj.cell(row = 1, column = i)
        if cell_obj.value == "Pr.kodas" or cell_obj.value == "Kodas":
            column_of_sku = i
        elif cell_obj.value == "Pr.pavad." or cell_obj.value == "Pavadinimas":
            column_of_title = i
        elif cell_obj.value == "Sand.likutis" or cell_obj.value == "Kiekis":
            column_of_quantity = i
    
    template_wb_obj = openpyxl.load_workbook("templates/unknown_stock_example.xlsx")
    excel_date = str(datetime.now().strftime("%Y-%m-%d"))
    file_date = today_date[5] + today_date[6] + today_date[8] + today_date[9] 
    template_sheet_obj = template_wb_obj.active
    insert_column_values(False, None, m_row - 1, None, template_sheet_obj, 2, "center", excel_date, False, False, None, None, False, stock_row, manufacturers)
    insert_column_values(True, sheet_obj, m_row, column_of_sku, template_sheet_obj, 3, "left", None, False, False, None, None, False, stock_row, manufacturers)
    insert_column_values(True, sheet_obj, m_row, column_of_title, template_sheet_obj, 5, "left", None, False, True, 4, "left", False, stock_row, manufacturers)
    insert_column_values(True, sheet_obj, m_row, column_of_quantity, template_sheet_obj, 6, "center", None, True, False, None, None, False, stock_row, manufacturers)

    stock_row = m_row - 2
    for i in range(2,  stock_row):
        template_cell_obj = template_sheet_obj.cell(row = i, column = 5)
        if isinstance(template_cell_obj.value, str):
            if "rx" in template_cell_obj.value.lower() or "radeon" in template_cell_obj.value.lower():
                template_sheet_obj.delete_rows(i)
                stock_row = stock_row - 1

    template_wb_obj.save(report_path + "\\" + "NVD-YYY-INV_kilobaitas-" +  file_date + ".xlsx")

def create_unknown_sales_report(manufacturers, export_path, prior_week_start, prior_week_end, report_path, sales_row):
    try:
        options = Options()
        options.add_experimental_option("detach", True)
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(executable_path="chromedriver.exe", options=options)
        driver.get("https://rekvizitai.vz.lt/imones/1/")
        driver.set_window_size(800, 800)
    except Exception:
        driver = None
        pass
    
    template_wb_obj = openpyxl.load_workbook("templates/unknown_sales_example.xlsx")
    template_sheet_obj = template_wb_obj.active

    vga_path = export_path + "\\" + "unknown_vga_sales.xlsx"
    vga_wb_obj = openpyxl.load_workbook(vga_path)
    vga_sheet_obj = vga_wb_obj.active
    max_col = vga_sheet_obj.max_column
    m_row = vga_sheet_obj.max_row
    for i in range(1, max_col + 1):
        cell_obj = vga_sheet_obj.cell(row = 1, column = i)
        if cell_obj.value == "Vertybes kodas" or cell_obj.value == "Vertybės kodas":
            column_of_sku = i
        elif cell_obj.value == "Vertybes pavadinimas" or cell_obj.value == "Vertybės pavadinimas":
            column_of_title = i
        elif cell_obj.value == "Kiekis":
            column_of_quantity = i
        elif cell_obj.value == "Op.data":
            column_of_date = i
        elif cell_obj.value == "Op.Nr.":
            column_of_opnum = i
        elif cell_obj.value == "BE PVM (val.)":
            column_of_price = i
        elif cell_obj.value == "Part.Im.kodas" or cell_obj.value == "Part.Įm.kodas":
            column_of_customer = i

    insert_column_values(True, vga_sheet_obj, m_row, column_of_sku, template_sheet_obj, 2, "left", None, False, False, None, None, False, sales_row, manufacturers)
    insert_column_values(True, vga_sheet_obj, m_row, column_of_title, template_sheet_obj, 4, "left", None, False, True, 3, "left", False, sales_row, manufacturers)
    insert_column_values(True, vga_sheet_obj, m_row, column_of_quantity, template_sheet_obj, 5, "center", None, True, False, None, None, False, sales_row, manufacturers)
    insert_column_values(True, vga_sheet_obj, m_row, column_of_date, template_sheet_obj, 6, "center", None, False, False, None, None, True, sales_row, manufacturers)
    insert_column_values(True, vga_sheet_obj, m_row, column_of_opnum, template_sheet_obj, 7, "center", None, True, False, None, None, False, sales_row, manufacturers)
    insert_column_values(True, vga_sheet_obj, m_row, column_of_price, template_sheet_obj, 9, "right", None, True, False, None, None, False, sales_row, manufacturers)
    insert_column_values(False, None, m_row - 1, None, template_sheet_obj, 10, "left", "EUR", False, False, None, None, False, sales_row, manufacturers)
    scrap_customer_info(vga_sheet_obj, m_row - 2, column_of_customer, template_sheet_obj, 12, "left", "LT-54447", driver, sales_row)
    insert_column_values(False, None, m_row - 1, None, template_sheet_obj, 13, "left", "LT", False, False, None, None, False, sales_row, manufacturers)
    template_sheet_obj.delete_rows(m_row - 1)
    sales_row = m_row - 2

    nb_path = export_path + "\\" + "unknown_nb_sales.xlsx"
    nb_wb_obj = openpyxl.load_workbook(nb_path)
    nb_sheet_obj = nb_wb_obj.active
    max_col = nb_sheet_obj.max_column
    m_row = nb_sheet_obj.max_row

    insert_column_values(True, nb_sheet_obj, m_row, column_of_sku, template_sheet_obj, 2, "left", None, False, False, None, None, False, sales_row, manufacturers)
    insert_column_values(True, nb_sheet_obj, m_row, column_of_title, template_sheet_obj, 4, "left", None, False, True, 3, "left", False, sales_row, manufacturers)
    insert_column_values(True, nb_sheet_obj, m_row, column_of_quantity, template_sheet_obj, 5, "center", None, True, False, None, None, False, sales_row, manufacturers)
    insert_column_values(True, nb_sheet_obj, m_row, column_of_date, template_sheet_obj, 6, "center", None, False, False, None, None, True, sales_row, manufacturers)
    insert_column_values(True, nb_sheet_obj, m_row, column_of_opnum, template_sheet_obj, 7, "center", None, True, False, None, None, False, sales_row, manufacturers)
    insert_column_values(True, nb_sheet_obj, m_row, column_of_price, template_sheet_obj, 9, "right", None, True, False, None, None, False, sales_row, manufacturers)
    insert_column_values(False, None, m_row - 1, None, template_sheet_obj, 10, "left", "EUR", False, False, None, None, False, sales_row, manufacturers)
    scrap_customer_info(nb_sheet_obj, m_row - 1, column_of_customer, template_sheet_obj, 12, "left", "LT-54447", driver, sales_row)
    insert_column_values(False, None, m_row - 1, None, template_sheet_obj, 13, "left", "LT", False, False, None, None, False, sales_row, manufacturers)
    sales_row = sales_row + m_row - 1

    export_list = []
    for path in os.listdir(export_path):
        if os.path.isfile(os.path.join(export_path, path)):
            export_list.append(path)
    print(export_list)

    for export in export_list:
        if "_zal.xlsx" in export:
            zal_path = export_path + "\\" + export
            zal_wb_obj = openpyxl.load_workbook(zal_path)
            zal_sheet_obj = zal_wb_obj.active
            max_col = zal_sheet_obj.max_column
            m_row = zal_sheet_obj.max_row
            for i in range(1, max_col + 1):
                cell_obj = zal_sheet_obj.cell(row = 1, column = i)
                if cell_obj.value == "ˇaliava" or cell_obj.value == "Žaliava":
                    column_of_sku = i
                elif cell_obj.value == "ˇ.pavadinimas" or cell_obj.value == "Ž.pavadinimas":
                    column_of_title = i
                elif cell_obj.value == "ˇ.kiekis" or cell_obj.value == "Ž.kiekis":
                    column_of_quantity = i
                elif cell_obj.value == "Data":
                    column_of_date = i
                elif cell_obj.value == "Op.nr.":
                    column_of_opnum = i
                
            insert_column_values(True, zal_sheet_obj, m_row, column_of_sku, template_sheet_obj, 2, "left", None, False, False, None, None, False, sales_row, manufacturers)
            insert_column_values(True, zal_sheet_obj, m_row, column_of_title, template_sheet_obj, 4, "left", None, False, True, 3, "left", False, sales_row, manufacturers)
            insert_column_values(True, zal_sheet_obj, m_row, column_of_quantity, template_sheet_obj, 5, "center", None, True, False, None, None, False, sales_row, manufacturers)
            insert_column_values(True, zal_sheet_obj, m_row, column_of_date, template_sheet_obj, 6, "center", None, False, False, None, None, True, sales_row, manufacturers)
            insert_column_values(True, zal_sheet_obj, m_row, column_of_opnum, template_sheet_obj, 7, "center", None, True, False, None, None, False, sales_row, manufacturers)
            insert_column_values(False, None, m_row - 1, None, template_sheet_obj, 11, "left", "PC Production", False, False, None, None, False, sales_row, manufacturers)
            sales_row = sales_row + m_row - 2

    
    for i in range(2,  sales_row):
        template_cell_obj = template_sheet_obj.cell(row = i, column = 4)
        if isinstance(template_cell_obj.value, str):
            if "rx" in template_cell_obj.value.lower() or "radeon" in template_cell_obj.value.lower():
                template_sheet_obj.delete_rows(i)
                sales_row = sales_row - 1
            

    file_date = prior_week_start[5] + prior_week_start[6] + prior_week_start[8] + prior_week_start[9] + "-" + prior_week_end[5] + prior_week_end[6] + prior_week_end[8] + prior_week_end[9]
    template_wb_obj.save(report_path + "\\" + "NVD-YYY-POS_kilobaitas-" +  file_date + ".xlsx")
    try:
        driver.quit()
    except Exception:
        pass

def main(ui_self, user_name, user_password):
    # Export data
    export_list = []
    filter_list = []
    f = open("data/ExportSettings.json",)
    f_data = json.load(f)
    for export in f_data:
        temp_list = []
        export_list.append(export)
        for filter in f_data[export]:
            temp_list.append(filter)
        filter_list.append(temp_list)
    f.close()

    # Manufacturer data
    manufacturers = []
    f = open("data/ManufacturerData.json",)
    f_data = json.load(f)
    for manufacturer in f_data["manufacturers"]:
        manufacturers.append(manufacturer)
    f.close()

    # Variables - Main variables
    unknown_sales_row = 0
    unknown_stock_row = 0

    # Variables - Date and Time variables
    today_date = datetime.now().strftime("%Y.%m.%d")
    prior_week_end = (datetime.now() - timedelta(days=((datetime.now().isoweekday()) % 7))).strftime("%Y.%m.%d")
    prior_week_start = ((datetime.now() - timedelta(days=((datetime.now().isoweekday()) % 7))) - timedelta(days=6)).strftime("%Y.%m.%d")
    translate_date_begin = prior_week_start[5] + prior_week_start[6] + prior_week_start[8] + prior_week_start[9] + prior_week_start[0] + prior_week_start[1] + prior_week_start[2] + prior_week_start[3]
    translate_date_end = prior_week_end[5] + prior_week_end[6] + prior_week_end[8] + prior_week_end[9] + prior_week_end[0] + prior_week_end[1] + prior_week_end[2] + prior_week_end[3]
    translate_date_begin_lt = ((datetime.now() - timedelta(days=((datetime.now().isoweekday()) % 7))) - timedelta(days=6)).strftime("%Y-%m-%d")
    translate_date_end_lt = (datetime.now() - timedelta(days=((datetime.now().isoweekday()) % 7))).strftime("%Y-%m-%d")

    # Variables - Variables for files and directories
    default_path = "C:/Dokumentai/Automatiškai sugeneruoti duomenys"
    f = open("data/DirectoryData.json", encoding='utf-8')
    f_data = json.load(f)
    directories = f_data["directories"]
    main_parent_directory = directories["report_directory"]
    f.close()
    main_directory_name = str(prior_week_start) + " - " + str(prior_week_end)
    main_path = os.path.join(main_parent_directory, main_directory_name)
    export_directory_name = "Eksportas iš įmonės programos"
    export_parent_directory = main_path
    export_path = os.path.join(export_parent_directory, export_directory_name)
    report_directory_name = "Savaitinės ataskaitos"
    report_parent_directory = main_path
    report_path = os.path.join(report_parent_directory, report_directory_name)

    try:
        # Step 1: Connects to the Application, creates directories and calibrates the main window to user monitor settings
        start_time = time.time()
        main_process, main_dlg = gf.connect(ui_self, user_name, user_password)
        
        if default_path == main_parent_directory and not os.path.isdir(default_path):
            os.makedirs(default_path)
        create_directories(ui_self, main_path, main_directory_name, export_path, report_path, export_directory_name, report_directory_name)

        # gf.print_ui("Tikrinama ChromeDriver versija. Prašome palaukti...", ui_self)
        # current_directory = os.getcwd()
        # check_driver(current_directory)
        # gf.print_ui("ChromeDriver versija yra tinkama tolimesniam darbui.", ui_self)

        pywinauto.mouse.move(coords=(10,10))
        time.sleep(1)

        gf.close_window(ui_self, "TFormEinamiejiVertybiuLikuciai", main_dlg)
        gf.close_window(ui_self, "TFormPrekiuPardavimaiPagalPartnerius", main_dlg)
        gf.close_window(ui_self, "TFormZaliavuPanaudojimasGaminiuose", main_dlg)

        scale = str(gf.find_current_dpi(ui_self, "meniu.png", 0.8))
        
        gf.print_ui("Tikrinama įmonės programos versija...", ui_self)
        time.sleep(2)
        try:
            window_text = []
            for child in main_dlg.children():
                window_text += child.texts()
                
            version = None
            if "\tv. 4.0.48.248" in window_text:
                gf.print_ui("Aptikta v. 4.0.48.248 įmonės programos versija. Versija palaikoma.", ui_self)
                version = "v. 4.0.48.248"
            elif "\tv. 4.0.48.243" in window_text:
                gf.print_ui("Aptikta v. 4.0.48.243 įmonės programos versija. Versija palaikoma.", ui_self)
                version = "v. 4.0.48.243"
            else:
                gf.print_ui("Įmonės programos versija nėra palaikoma, darbai bus atliekami kaip v. 4.0.48.248 versijai.", ui_self)
        except Exception as e: gf.print_ui(e, ui_self)

        # Step 2: Exports stock from the program
        gf.print_ui("Pradedama eksportuoti likučius iš įmonės programos...", ui_self)
        gf.navigate_menu(ui_self, scale + "ataskaitos.png", 0.8, scale + "vertybiu-likuciai.png", 0.8, scale + "faktiniai-likuciai.png", 0.8, main_dlg)
        time.sleep(3)
        stock_dlg = main_dlg.child_window(class_name="TFormEinamiejiVertybiuLikuciai")
        stock_dlg.set_focus()
        time.sleep(2)
        stock_dlg.minimize()
        stock_dlg.maximize()
        gf.refresh_data(scale)
        for idx, export in enumerate(export_list):
            stock_export(export, filter_list[idx][0], filter_list[idx][1], filter_list[idx][2], main_process, export_path, scale)
            gf.print_ui(export + " likučiai eksportuoti...", ui_self)
        stock_export("unknown", "kk/vga", "", "", main_process, export_path, scale)
        gf.print_ui("Unkown likučiai eksportuoti...", ui_self)
        stock_dlg.close()
        gf.print_ui("Pabaigta eksportuoti likučius.", ui_self)

        # Step 3: Exports sales from the program
        gf.print_ui("Pradedama eksportuoti pardavimus iš įmonės programos...", ui_self)

        gf.navigate_menu(ui_self, scale + "ataskaitos.png", 0.8, scale + "prekybos-analize.png", 0.9, scale + "prekiu-pardavimai.png", 0.9, main_dlg)
        sales_dlg = main_dlg.child_window(class_name="TFormPrekiuPardavimaiPagalPartnerius")
        sales_dlg.set_focus()
        time.sleep(2)
        sales_dlg.minimize()
        sales_dlg.maximize()

        if version == "v. 4.0.48.243":
            textbox_start = sales_dlg.Edit11
            textbox_end = sales_dlg.Edit10
        else:
            textbox_start = sales_dlg.Edit13
            textbox_end = sales_dlg.Edit12

        combobox_order = sales_dlg.child_window(title="Pagal partnerius", class_name="TRxDBLookupCombo")
        combobox_type = sales_dlg.child_window(best_match="Sumin")

        gf.input_date_by_locale(textbox_start, textbox_end, translate_date_begin, translate_date_end, translate_date_begin_lt, translate_date_end_lt)

        time.sleep(1)
        combobox_order.click()
        time.sleep(1)
        check_if_by_group(ui_self, scale)
        combobox_type.click()
        time.sleep(1)
        gf.locate_and_click(scale + "detali-sum-prekes.png", 0.8)
        time.sleep(1)
        gf.refresh_data(scale)
        for idx, export in enumerate(export_list):
            sales_export(export, filter_list[idx][0], filter_list[idx][1], filter_list[idx][2], main_process, export_path, scale)
            gf.print_ui(export + " pardavimai eksportuoti...", ui_self)

        combobox_type = sales_dlg.child_window(best_match="Detali+sum")
        combobox_type.click()
        time.sleep(1)
        gf.locate_and_click(scale + "detali-plius-partijos.png", 0.8)
        time.sleep(1)
        gf.refresh_data(scale)
        sales_export("unknown_vga", "kk/vga", "", "", main_process, export_path, scale)
        gf.print_ui("Unkown VGA pardavimai eksportuoti...", ui_self)
        sales_export("unknown_nb", "k/nb_", "", "tx", main_process, export_path, scale)
        gf.print_ui("Unkown NB pardavimai eksportuoti...", ui_self)
        sales_dlg.close()
        gf.print_ui("Pabaigta eksportuoti pardavimus.", ui_self)

        # Step 4: Exports PC sales (Žaliavos) from the program
        gf.print_ui("Pradedama eksportuoti žaliavas iš įmonės programos...", ui_self)
        gf.navigate_menu(ui_self, scale + "ataskaitos.png", 0.8, scale + "gamybos-analize.png", 0.9, scale + "zaliavu-panaudojimas-gaminiuose.png", 0.8, main_dlg)
        pc_sales_dlg = main_dlg.child_window(class_name="TFormZaliavuPanaudojimasGaminiuose")
        pc_sales_dlg.set_focus()
        time.sleep(2)
        pc_sales_dlg.minimize()
        pc_sales_dlg.maximize()
        textbox_start = pc_sales_dlg.Edit14
        textbox_end = pc_sales_dlg.Edit13
        combobox_type = pc_sales_dlg.Combobox
        combobox_group = pc_sales_dlg.TRxDBLookupCombo2

        gf.input_date_by_locale(textbox_start, textbox_end, translate_date_begin, translate_date_end, translate_date_begin_lt, translate_date_end_lt)

        time.sleep(1)
        combobox_type.click()
        time.sleep(1)
        gf.locate_and_click(scale + "detali.png", 0.8)
        time.sleep(1)
        combobox_group.click()
        time.sleep(1)
        gf.locate_and_click(scale + "pagal-gaminius.png", 0.8)
        time.sleep(1)
        gf.refresh_data(scale)
        pc_sales_export("unknown_gt730", "", "730", main_process, export_path, scale)
        gf.print_ui("Unkown GT 730 žaliavos eksportuotos...", ui_self)
        pc_sales_export("unknown_gt1030", "", "1030", main_process, export_path, scale)
        gf.print_ui("Unkown GT 1030 žaliavos eksportuotos...", ui_self)
        pc_sales_export("unknown_gtx", "", "gtx", main_process, export_path, scale)
        gf.print_ui("Unkown GTX žaliavos eksportuotos...", ui_self)
        pc_sales_export("unknown_rtx", "", "rtx", main_process, export_path, scale)
        gf.print_ui("Unkown RTX žaliavos eksportuotos...", ui_self)
        pc_sales_dlg.close()
        gf.print_ui("Pabaigta eksportuoti žaliavas.", ui_self)
        gf.print_ui("Visi duomenys iš įmoonės programos eksportuoti sėkmingai.", ui_self)

        # Step 5: Creates reports from exported data
        gf.print_ui("Pradedama ruošti ataskaitas...", ui_self)
        for export in export_list:
            create_report(export, "sales", manufacturers, export_path, today_date, prior_week_start, prior_week_end, report_path)
            create_report(export, "stock", manufacturers, export_path, today_date, prior_week_start, prior_week_end, report_path)
            gf.print_ui(export + " pardavimų ir likučių ataskaitos suformuotos...", ui_self)
        create_unknown_stock_report(manufacturers, export_path, today_date, report_path, unknown_stock_row)
        gf.print_ui("Unkown likučių ataskaita suformuota...", ui_self)
        create_unknown_sales_report(manufacturers, export_path, prior_week_start, prior_week_end, report_path, unknown_sales_row)
        gf.print_ui("Unkown pardavimų ataskaita suformuota...", ui_self)
        gf.print_ui("Pabaigta ruošti ataskaitas.", ui_self)
        end_time = time.time()
        seconds_spent_executing = end_time - start_time
        executing_time = timedelta(seconds=seconds_spent_executing)
        gf.print_ui("Visi savaitinių ataskaitų modulio darbai pabaigti per:\n [" + str(executing_time) + "]", ui_self)
    except Exception as e:
        gf.print_ui("DĖMESIO! Įvyko klaida, programa negali toliau tęsti darbo. Klaidos pranešimas: " + str(e), ui_self)

if __name__ == "__main__":
    main(None, "", "")
