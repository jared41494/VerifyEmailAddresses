import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.common.exceptions import NoSuchElementException


def check_exists_by_xpath(driver, xpath):
    try:
        driver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True


excelDirPath = "Excel"
excelFilePath = os.path.join(excelDirPath, "NSW.xlsx")
successExcelFP = os.path.join(excelDirPath, "SuccessEmails.xlsx")
failedExcelFP = os.path.join(excelDirPath, "FailedEmails.xlsx")

print("READING EXCEL {0}...".format(excelFilePath))
df = pd.read_excel(excelFilePath)
counter = 0
print("OPENING CHROME...")
driver = webdriver.Chrome(executable_path=r"D:\chromedriver\chromedriver.exe")
driver.get('https://www.verifyemailaddress.org/')
mainWindow = driver.window_handles[0]
print("PROCESSING EMAILS...")
successDF = {}
failedDF = {}
successDF['EMAIL'] = []
failedDF['EMAIL'] = []
for i in df.index:
    try:
        if counter == 100:  # TESTING
            break
        time.sleep(1)
        email = df['EMAIL'][i]
        print("VERIFYING EMAIL: {0}".format(email))
        driver.find_element(
            By.XPATH, "/html/body/header/form/fieldset/input").send_keys(email)

        if check_exists_by_xpath(driver, "/html/body/header/form/fieldset/div/div/div/iframe"):
            iframeElement = driver.find_element(
                By.XPATH, "/html/body/header/form/fieldset/div/div/div/iframe")
            driver.switch_to.frame(iframeElement)
            driver.find_element(
                By.XPATH, '//*[@id="recaptcha-anchor"]').click()
            time.sleep(10)
            driver.switch_to.window(mainWindow)

        driver.find_element(
            By.XPATH, "/html/body/header/form/fieldset/button").click()
        time.sleep(2)
        successEle = driver.find_elements(By.CSS_SELECTOR, "li.success")
        failureEle = driver.find_elements(By.CSS_SELECTOR, "li.failure")
        successCount = len(successEle)
        failureCount = len(failureEle)
        if failureCount > 0:
            failedDF['EMAIL'].append(email)
        else:
            successDF['EMAIL'].append(email)
        driver.find_element(By.XPATH, '//*[@id="result"]/nav/a[2]').click()
        counter += 1
    except Exception as e:
        print(e)

print("WRITING EXCEL {0}...".format(successExcelFP))
df = pd.DataFrame(successDF)
df.to_excel(successExcelFP)

print("WRITING EXCEL {0}...".format(failedExcelFP))
df = pd.DataFrame(failedDF)
df.to_excel(failedExcelFP)

print("PROCESS COMPLETE")
