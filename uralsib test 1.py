from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
# import win32com.client
import pandas as pd
import urllib
import time

# Excel = win32com.client.Dispatch("Excel.Application")
# Excel.Visible=True
# wb =  Excel.Workbooks.Open(r'C:\Users\79777\Desktop\Примеры.xlsx')
# sheet = wb.ActiveSheet
# vals = [r[0].value for r in sheet.UsedRange]
# k=4
# for i in range(4, len(vals), 4):
#    name = print(vals[i])
#    secondname = print(vals[i+1])
#    surname = print(vals[i+2])
#    dateofbirth = print(vals[i+3])
#    k+=4
# print (ex.at[0,"Имя"])
# print(ex.index[-1])
ex = pd.read_excel(r'C:\Users\79777\Desktop\Примеры.xlsx')
driver = webdriver.Chrome()
driver.implicitly_wait(20)
driver.get("http://fssprus.ru/")


driver.maximize_window()
assert "Федеральная" in driver.title
driver.find_element(By.XPATH,"//button[@class='tingle-modal__close']").click()
driver.find_element(By.XPATH,"//a[contains(text(), 'Расширенный') and @class='btn btn-light']").click()
for i in range (0, ex.index[-1]):
    driver.find_element(By.XPATH,"//*[@placeholder='Фамилия']").send_keys(ex.at[i, "Фамилия"])
    driver.find_element(By.XPATH,"//*[@placeholder='Имя']").send_keys(ex.at[i, "Имя"])
    driver.find_element(By.XPATH,"//*[@placeholder='Отчество']").send_keys(ex.at[i, "Отчество"])
    driver.find_element(By.XPATH,"//*[@placeholder='дд.мм.ггг']").send_keys(ex.at[i, "Дата рождения"].strftime("%d-%m-%Y"))
    driver.find_element(By.XPATH,"//*[text()='Узнай о своих долгах']").click()
    driver.find_element(By.XPATH,"//div[contains(@class, 'main')]//button[@type='submit']").click()
    n=1
    time.sleep(10)
    while (n<40000):
        
        filename = f"D:\Капчи\{n}_capcha.png" 
        with open(filename, 'wb') as file:
#identify image to be captured
#write fil
            img = driver.find_element(By.XPATH, "//img[@id='capchaVisual']")
            file.write(img.screenshot_as_png)
            img.click()
            time.sleep(10)
        n+=1
driver.close()