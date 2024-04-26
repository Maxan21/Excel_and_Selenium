import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
file ="" #Расположение таблицы Excel
xl = pd.read_excel(file)
xl1=load_workbook(file)
s=xl1['Лист1']
o = Options()
o.add_experimental_option("detach", True)
browser = webdriver.Chrome(options=o)
url='https://egrul.nalog.ru/index.html'
browser.get(url)
for i in range(xl.shape[0]):
    b = "B" + str(i + 2)
    if xl['состояние'][i]==1:
        n=0
        while n<=3:
            if n==3:
                browser.refresh()
                input_1 = browser.find_element(By.ID, "query")
                input_1.send_keys("0"+str(int(xl['инн'][i])))
                button_element = browser.find_element(By.ID, 'btnSearch')
                button_element.click()
                time.sleep(1.5)
            try:
                if n<3:
                    input_1 = browser.find_element(By.ID, "query")
                    input_1.send_keys(str(int(xl['инн'][i])))
                    button_element = browser.find_element(By.ID, 'btnSearch')
                    button_element.click()
                    time.sleep(1.5)
                result = browser.find_element(By.CLASS_NAME, "res-text")
                if "Дата прекращения деятельности" in result.text:
                    s[b] = 'ликвидирована'
                elif "Дата прекращения деятельности" not in result.text and "ЛИКВИДАТОР" in result.text:
                    s[b]="В очереди на ликвидацию"
                else:
                    s[b] = 'существует'
                input_1.clear()
                break
            except Exception as e:
                s[b] = 'компании не существует или не получилось обработать запрос'
                input_1.clear()
            n+=1
            browser.refresh()
        xl1.save(file)


xl1.close()



