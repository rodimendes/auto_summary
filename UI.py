from tkinter import *
from datetime import datetime
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui
import time


# OPEN AND READ DOC TO INSERT
def submit_summary():
    pyautogui.PAUSE = 2
    data = data_entry.get()
    texto_sumario = sumario_text.get('1.0', END)

    # INSERT DATE TO SUMMARY
    summary_date = data
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
    # time.sleep(2)
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
    workbook = load_workbook("/Users/rodrigocamila/PycharmProjects/ajuda-a-esposa/registro-sumario.xlsx")

    worksheet = workbook.active

    date_column = worksheet['B']

    for cell in date_column:
        if cell.value == formatted_date:
            line = cell.row
            worksheet[f'H{line}'] = texto_sumario

    workbook.save(filename="/Users/rodrigocamila/PycharmProjects/ajuda-a-esposa/registro-sumario.xlsx")


def clear_fields():
    data_entry.delete(0, END)
    sumario_text.delete('1.0', END)



today = datetime.now().strftime("%d/%m/%Y")

window = Tk()
window.title("Sumários para a Camila")

avatar = PhotoImage(file='/Users/rodrigocamila/Learning Programming/Portfolio_projects/ajuda-a-esposa/mila_avatar_sm.png')
canvas = Canvas(height=300, width=500)
canvas.create_image(250, 150, image=avatar)
canvas.grid(column=0, row=0, columnspan=2)

# Labels
data_label = Label(text="Data para cadastro:")
data_label.grid(column=0, row=1)
sumario_label = Label(text="Sumário:")
sumario_label.grid(column=0, row=2)

# Entries
data_entry = Entry(width=20)
data_entry.insert(END, string=today)
data_entry.grid(column=1, row=1)
sumario_text = Text(height=10, width=50)
sumario_text.focus()
sumario_text.grid(column=1, row=2, padx=5, pady=5)

# Buttons
clear_button = Button(text='Limpar campos', width=15, command=clear_fields)
clear_button.grid(column=0, row=3, padx=5, pady=10)
submit_button = Button(text='Submeter sumário', width=15, command=submit_summary)
submit_button.grid(column=1, row=3, columnspan=2, padx=5, pady=10)

window.mainloop()
