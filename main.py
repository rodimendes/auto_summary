from flask import Flask, render_template, request
import datetime
import pyautogui
import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By

application = Flask(__name__)

@application.route('/', methods=['GET', 'POST'])
def fill_summary():
    if request.method == 'POST':
        pyautogui.PAUSE = 2
        data = request.form

        # INSERT DATE TO SUMMARY
        summary_date = data['sum-date']
        formatted_date = datetime.datetime.strptime(summary_date, '%Y-%m-%d')
        # print((summary_date), (formatted_date), data['summary'])
        
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
        where_to_write.send_keys(data['summary'])
        time.sleep(3)
        driver.switch_to.default_content()
        time.sleep(2)
        submit = driver.find_element(By.XPATH, '//*[@id="SummaryFirstGradeDetailsManagerHolder"]/div/div[5]/input[2]')
        submit.click()
        time.sleep(2)
        pyautogui.press('enter')
        # time.sleep(2)
        driver.quit()

        if weekday == 2:
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
            fill_summary_text.send_keys(data['summary'])
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
                worksheet[f'H{line}'] = data['summary']

        workbook.save(filename="/Users/rodrigocamila/PycharmProjects/ajuda-a-esposa/registro-sumario.xlsx")
        
        return render_template('form.html')
        # return '<h1>Até aqui está indo bem</h1>'

    return render_template('form.html')
        
    
    
if __name__ == '__main__':
    application.run(debug=True)