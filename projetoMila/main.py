from datetime import datetime
from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import pyautogui
import time

# MOUSE ARROW POSITION
# time.sleep(5)
# print(pyautogui.position())

# OPEN AND READ DOC TO INSERT
def submit(data, texto):
    texto_sumario = texto
    print(texto_sumario)
    # with open("projetoMila/teste-sumario.txt", "r") as sumario:
    #     texto = sumario.read()

    # INSERT DATE TO SUMMARY
    summary_date = data
    # summary_date = input("Insira a data (dd/mm/aaaa): ")
    formatted_date = datetime.strptime(summary_date, '%d/%m/%Y')

    # MAIN PROGRAM
    while formatted_date.isoweekday() > 5:
        print("A data coincide com final de semana.")
        summary_date = input("Insira a data (dd/mm/aaaa): ")
        formatted_date = datetime.strptime(summary_date, '%d/%m/%Y')

    weekday = formatted_date.isoweekday()
    weekdays_position = [[260, 610], [380, 610], [500, 610], [620, 610], [740, 610]]
    x = weekdays_position[weekday - 1][0]
    y = weekdays_position[weekday - 1][1]


    # ACCESS HOMEPAGE TO INTERACT
    site_sei = "https://siga1.edubox.pt/SEI/autentication.aspx"
    site_inovar = 'https://inovar.aeandresoares.pt/InovarAlunos/Inicial.wgx'
    chrome_driver_path = "/Users/rodrigocamila/PycharmProjects/chromedriver"

    driver = webdriver.Chrome(executable_path=chrome_driver_path)
    driver.get(url=site_sei)

    driver.maximize_window()

    user = driver.find_element(By.ID, 'ContentPlaceHolder1_username')
    user.send_keys('BRG.PF2261758')
    # time.sleep(2)
    password = driver.find_element(By.ID, 'ContentPlaceHolder1_password')
    password.send_keys('768457')
    time.sleep(2)
    enter_key = driver.find_element(By.ID, 'ContentPlaceHolder1_submit')
    enter_key.click()
    time.sleep(2)
    aulas = driver.find_element(By.ID, 'MASTER_MenuButton7')
    aulas.click()
    time.sleep(2)
    sumario = driver.find_element(By.ID, 'MASTER_SubMenuButton30')
    sumario.click()
    time.sleep(3)
    # verificar_data = driver.find_element(By.CLASS_NAME, 'hasDatepicker')
    # verificar_data.click()
    # time.sleep(2)
    pyautogui.click(x, y)
    time.sleep(5)
    driver.switch_to.frame(driver.find_element(By.TAG_NAME, "iframe"))
    where_to_write = driver.find_element(By.ID, 'tinymce')
    where_to_write.send_keys(texto_sumario)
    time.sleep(3)
    driver.switch_to.default_content()
    time.sleep(2)
    submit = driver.find_element(By.XPATH, '//*[@id="SummaryFirstGradeDetailsManagerHolder"]/div/div[5]/input[2]')
    submit.click()
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)
    
    driver.quit()

    if formatted_date.isoweekday() == 2:
        driver = webdriver.Chrome(executable_path=chrome_driver_path)
        driver.get(url=site_inovar)

        driver.maximize_window()

        user_inovar = driver.find_element(By.ID, 'TRG_29')
        user_inovar.send_keys('304185426')
        time.sleep(1)
        password_inovar = driver.find_element(By.ID, 'TRG_28')
        password_inovar.send_keys('304185426')
        time.sleep(2)
        enter_key_inovar = driver.find_element(By.ID, 'VWG_30')
        enter_key_inovar.click()
        time.sleep(3)
        teacher_area = driver.find_element(By.ID, 'VWG_53')
        teacher_area.click()
        time.sleep(2)
        summary_button = driver.find_element(By.ID, 'VWG_105')
        summary_button.click()
        time.sleep(2)
        chosen_day = driver.find_element(By.ID, 'VWG_143_E1')
        chosen_day.click()
        time.sleep(2)
        fill_summary_text = driver.find_element(By.ID, 'TRG_433')
        fill_summary_text.send_keys(texto_sumario)
        time.sleep(2)
        submit_inova = driver.find_element(By.ID, 'VWG_412')
        submit_inova.click() 

        time.sleep(5)

        driver.quit()


    # FILL EXCEL FILE
    # GET ACCESS AND EDIT ESPECIFIC CELLS
    workbook = load_workbook("projetoMila/registro-sumario.xlsx")

    worksheet = workbook.active

    date_column = worksheet['B']

    # date_cells = date_column[3:]

    for cell in date_column:
        if cell.value == formatted_date:
            line = cell.row
            worksheet[f'H{line}'] = texto_sumario
    #     # full_datetime = str(celula.value)
    #     # date = full_datetime.split(" ")[0]
    #     # print(date)
    #     # celula.value = texto_sumario

    today = datetime.now().strftime('%d_%m_%Y')

    workbook.save(filename=f"sumario_atualizado_{today}.xlsx")