# Andrew B. Reyes
# needs major OOP integration updated

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime
from pandas import ExcelWriter
from selenium.webdriver.common.action_chains import ActionChains
from Keys import *

df = pd.read_excel(open('Input_File.xlsx', 'rb'), sheet_name=0)

browser = webdriver.Chrome(executable_path="C:\\Users\\reyesa1\\Documents\\chromedriver.exe")
browser.get("https://totalpaycard.adp.com/paycardProxyOnline/jsp/landing.faces")

workingClient = []
clientStatus = []
updateStatus = []



def timeStamp(inputText):
    return (inputText + " - " + datetime.datetime.now().strftime("%Y/%m/%d") + " - " + datetime.datetime.now().strftime("%I:%M:%S %p"))
def login_initiation():
    username = input("Enter username: ")
    password = input("Enter password: ")
    try:
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located(
                (By.ID, "USER"))
        )
    except:
        print("Username input error occurred!")
        quit()
    finally:
        browser.find_element_by_id("USER").send_keys(username)
        browser.find_element_by_class_name("loginButton").click()
    try:
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located(
                (By.ID, "passwordForm:password"))
        )
    except:
        print("Passphrase input error occurred!")
        quit()
    finally:
        browser.find_element_by_id("passwordForm:password").send_keys(password)
        browser.find_element_by_class_name("loginButton").click()
    print("Successful Login...")
def navi_VieworUpdate_Client():
    WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="revit_navigation_NavItem_3_label"]'))
    )

    element_to_hover_over = browser.find_element_by_id("revit_navigation_NavHoverItem_0_label")

    hover = ActionChains(browser).move_to_element(element_to_hover_over)
    hover.perform()
    WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.ID, 'revit_navigation_NavItem_3_label'))
    )
    browser.find_element_by_id('revit_navigation_NavItem_3_label').click()
def newLocationsearch(locationName):
    WebDriverWait(browser, 3).until(
        EC.presence_of_element_located(
            (By.ID, "txtISIId"))
    )
    browser.find_element_by_id('txtISIId').send_keys(locationName)
    browser.find_element_by_id('btnSearch').click()
    try:
        WebDriverWait(browser, 2).until(
            EC.presence_of_element_located(
                (By.ID, "clientSearchResultsGrid_row_0_cell_4_action"))
        )
        browser.find_element_by_id('clientSearchResultsGrid_row_0_cell_4_action').click()
        browser.find_element_by_link_text("Edit Detail").click()
        clientStatus.append("Client Found")
    except:
        clientStatus.append("Client Not Found")
        pass
def client_portal_update(clientAccount_Name, clientAlias_ID):
    try:
        WebDriverWait(browser, 4).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="accAliasIdGrid_row_0_cell_0"]')))
        browser.find_element_by_xpath(
            "//*[@id='accAliasIdGrid_row_0_cell_0']/span/div/div[2]/input[contains(@id,'revit_form_ValidationTextBox_')]").clear()
        browser.find_element_by_xpath(
            "//*[@id='accAliasIdGrid_row_0_cell_0']/span/div/div[2]/input[contains(@id,'revit_form_ValidationTextBox_')]").send_keys(
            clientAccount_Name)

        browser.find_element_by_xpath(
            "//*[@id='accAliasIdGrid_row_0_cell_1']/span/div/div[2]/input[contains(@id,'revit_form_ValidationTextBox_')]").clear()
        browser.find_element_by_xpath(
            "//*[@id='accAliasIdGrid_row_0_cell_1']/span/div/div[2]/input[contains(@id,'revit_form_ValidationTextBox_')]").send_keys(
            clientAlias_ID)
        browser.find_element_by_id('addClientButton').click()
    except:
        pass
    try:
        WebDriverWait(browser, 1).until(
            EC.presence_of_element_located(
                (By.ID, 'confirmSubclientNodeIdsDialogOk')))
        browser.find_element_by_id('confirmSubclientNodeIdsDialogOk').click()
    except:
        pass
    try:
        WebDriverWait(browser, 4).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(), ' has been updated successfully')]")))
        updateStatus.append("Successful Update")
    except:
        updateStatus.append("unSuccessful Update")
        pass

login_initiation()
print(timeStamp("First Login Successful - Working on " + str(len(df))))
navi_VieworUpdate_Client()
ctr = 0

for index, row in df.iterrows():
    ctr += 1

    clientName = (row['ISI_ID'])
    workingClient.append(clientName)
    clientAccount = (row['Funding_Account'])
    clientAlias = (row['Alias_ID'])

    newLocationsearch(clientName)
    client_portal_update(clientAccount, clientAlias)

    navi_VieworUpdate_Client()
    try:
        WebDriverWait(browser, 1).until(
            EC.presence_of_element_located(
                (By.ID, 'revit_navigation_NavItem_3_label')))
        browser.find_element_by_id('revit_navigation_NavItem_3_label').click()
    except:
        pass

    df2 = pd.DataFrame({
        "ISI ID": workingClient,
        "Client Status": clientStatus,
        "Update Status": updateStatus,
    })
    writer = ExcelWriter('Output_File.xlsx')
    df2.to_excel(writer, 'Sheet1', index=False)
    writer.save()

    # temp counter
    if ctr == 500 or ctr == 1000 or ctr == 2000 or ctr == 3000 or ctr == 4000 or ctr == 5000 or ctr == 6000:
        print(timeStamp(str(ctr)))
    else:
        pass

print(timeStamp("\nJob Complete for " + str(len(df))))
