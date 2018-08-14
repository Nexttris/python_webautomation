# Extract location funding account data from Visa Pas
# Andrew B. Reyes
import pandas as pd
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from myFunctions import clearSendkey
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import datetime
from myFunctions import *
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
import time
from selenium.webdriver.support.ui import Select
import pyautogui

df = pd.read_excel(open('Input_data.xlsx','rb'), sheet_name=0)

browser = webdriver.Chrome(executable_path="C:\\Users\\reyesa1\\Documents\\chromedriver.exe")
browser.get("https://www.visadpsprepaid.com/pas")


def relogin():
    try:
        browser.find_element_by_id('PleaseLogOn')
        loginScreen(browser)
        print(timeStamp("*****Relogin Initiated*****"))
    except:
        pass
def fundingAccount_Analysis():
    # GoTo funding accounts tab
    try:
        browser.find_element_by_id(
            'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
    except:
        pass
    # data collection code of funding accounts
    try:
        subclientid.append(location_ID)
        workingLocations.append(selectLocation)
        indexer.append("1")
        accountsList0.append(browser.find_element_by_id(
            'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl02_lblAccountNumber').text)
        expDate0.append(browser.find_element_by_id(
            'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl02_lblExpDate').text)
        primeStat0.append(browser.find_element_by_id(
            'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl02_lblPrimaryAccount').text)
        proxyId0.append(browser.find_element_by_xpath(
            '//*[@id="ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl02_ctl06_tblInfo"]/tbody/tr[1]/td[2]').get_attribute(
            'textContent').strip())
    except:
        accountsList0.append("N/A")
        expDate0.append("N/A")
        primeStat0.append("N/A")
        proxyId0.append("N/A")
        pass
    try:
        accountsList0.append(browser.find_element_by_id(
            'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl03_lblAccountNumber').text)
        expDate0.append(browser.find_element_by_id(
            'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl03_lblExpDate').text)
        primeStat0.append(browser.find_element_by_id(
            'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl03_lblPrimaryAccount').text)
        proxyId0.append(browser.find_element_by_xpath(
            '//*[@id="ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl03_ctl06_tblInfo"]/tbody/tr[1]/td[2]').get_attribute(
            'textContent').strip())
        subclientid.append(location_ID)
        workingLocations.append(selectLocation)
        indexer.append("2")
    except:
        pass
    try:
        accountsList0.append(browser.find_element_by_id(
            'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl04_lblAccountNumber').text)
        expDate0.append(browser.find_element_by_id(
            'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl04_lblExpDate').text)
        primeStat0.append(browser.find_element_by_id(
            'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl04_lblPrimaryAccount').text)
        proxyId0.append(browser.find_element_by_xpath(
            '//*[@id="ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl04_ctl06_tblInfo"]/tbody/tr[1]/td[2]').get_attribute(
            'textContent').strip())
        subclientid.append(location_ID)
        workingLocations.append(selectLocation)
        indexer.append("3")
    except:
        pass
def newLocationsearch():
    browser.get("https://www.visadpsprepaid.com/PAS/Location/LOCSearch.aspx?nav=1")
    try:
        clearSendkey(browser,
                     'ctl00_MainContentPanel_LOCSearchCriteriaControl1_LocationNameTextBox_LocationNameTextBox_InputTextBox',
                     selectLocation)
        browser.find_element_by_id('ctl00_MainContentPanel_LOCSearchCriteriaControl1_SearchLinkButton').click()
    except:
        pass
    try:
        # locate location and select after search
        try:
            browser.find_element_by_xpath('//*[@id="0_rowInd"]/td[1]/a').click()
            sub_id = browser.find_element_by_id('ctl00_BlueContentPanel_LOCTypeControl1_lblLocationIdValue').text
            if ''.join(sub_id.split()) == ''.join(location_ID.split()):
                pass
            else:
                browser.get("https://www.visadpsprepaid.com/PAS/Location/LOCSearch.aspx?nav=1")
                browser.find_element_by_xpath('//*[@id="1_rowInd"]/td[1]/a').click()
                sub_id = browser.find_element_by_id('ctl00_BlueContentPanel_LOCTypeControl1_lblLocationIdValue').text
                if ''.join(sub_id.split()) == ''.join(location_ID.split()):
                    pass
                else:
                    browser.get("https://www.visadpsprepaid.com/PAS/Location/LOCSearch.aspx?nav=1")
                    browser.find_element_by_xpath('//*[@id="2_rowInd"]/td[1]/a').click()
                    sub_id = browser.find_element_by_id('ctl00_BlueContentPanel_LOCTypeControl1_lblLocationIdValue').text
                    if ''.join(sub_id.split()) == ''.join(location_ID.split()):
                        pass
                    else:
                        pass
        except:
            pass
        # input function(s) here
        #add_funding_account("0000000002373370164", "9851", "1200 12th Ave S", "", "Seattle", "WA", "98144-2712")
        #GoTo funding accounts tab
        try:
            browser.find_element_by_id(
            'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
        except:
            pass
        set_FA_primary("9851")
        #expUpdate_addressUpdate("9025", "06", "28", "920 Winter St", "", "Waltham", "MA", "02451-1521")
        #set_FA_primary("9025")
        fundingAccount_Analysis()
    except:
        # appended info if location is not found
        indexer.append("0")
        subclientid.append(location_ID)
        workingLocations.append(selectLocation)
        accountsList0.append("N/A")
        expDate0.append("N/A")
        primeStat0.append("N/A")
        proxyId0.append("N/A")
        pass
def card_Inventory_Order():
    browser.get("https://www.visadpsprepaid.com/PAS/InventoryControl/ICCreateOrder.aspx")
    try:
        workingLocations.append(selectLocation)
        # adjust card program dropdown to Aline-MB
        WebDriverWait(browser, 20).until(
            EC.presence_of_element_located(
                (By.ID, "ctl00_BlueContentPanel_ICCardOrderHeader1_ICCardSelectorControl1_DropDownCardProgram_DropDownList"))
        )
        card_programDropdown = Select(browser.find_element_by_id("ctl00_BlueContentPanel_ICCardOrderHeader1_ICCardSelectorControl1_DropDownCardProgram_DropDownList"))
        card_programDropdown.select_by_index(9)

        clearSendkey(browser,
                     'ctl00_BlueContentPanel_ICCardOrderHeader1_ICCardSelectorControl1_DropDownLocation_SearchTextBox_SearchTextBox_InputTextBox',
                     selectLocation)

        time.sleep(2)
        location_select_Dropdown = Select(browser.find_element_by_id("ctl00_BlueContentPanel_ICCardOrderHeader1_ICCardSelectorControl1_DropDownLocation_DropDownList"))
        location_select_Dropdown.select_by_index(1)
    except:
        pass
    try:
        WebDriverWait(browser, 20).until(
            EC.presence_of_element_located(
                (By.ID, "ctl00_BlueContentPanel_ICCardOrderInformation1_DropDownCardDesign_DropDownList"))
        )
        card_design_Dropdown = Select(browser.find_element_by_id("ctl00_BlueContentPanel_ICCardOrderInformation1_DropDownCardDesign_DropDownList"))
        card_design_Dropdown.select_by_index(1)

        browser.find_element_by_id("ctl00_BlueContentPanel_ICCardOrderInformation1_TextBoxCardQuantity_TextBoxCardQuantity_InputTextBox").send_keys("25")
        browser.find_element_by_id("ctl00_BlueContentPanel_ICCardOrderInformation1_TextBoxComments").send_keys("Card order as request by R. Deshotel")
        browser.find_element_by_xpath("//*[text()[contains(.,'SUBMIT')]]").click()
        try:
            WebDriverWait(browser, 3).until(EC.alert_is_present(),
                                            'Timed out waiting for PA creation ' +
                                            'confirmation popup to appear.')
            alert = browser.switch_to.alert
            alert.accept()
        except:
            pass

        conf_code.append(browser.find_element_by_id("ctl00_BlueContentPanel_ICOrderQueueSearchCriteria1_TextBoxOrderNumber_TextBoxOrderNumber_InputTextBox").get_attribute("value"))
    except:
        conf_code.append("error")
        pass
def remove_funding_Account(account_number):
    try:
        browser.find_element_by_xpath("//*[text()[contains(.,'" + account_number + "')]]").click()
        browser.find_element_by_xpath("//*[text()[contains(.,'EDIT')]]").click()
        browser.find_element_by_xpath("//*[text()[contains(.,'DELETE ACCOUNT')]]").click()
    except:
        pass

    try:
        WebDriverWait(browser, 3).until(EC.alert_is_present(),
                                        'Timed out waiting for PA creation ' +
                                        'confirmation popup to appear.')
        alert = browser.switch_to.alert
        alert.accept()
    except:
        pass
def set_FA_primary(last_4):
    try:
        browser.find_element_by_xpath("//*[text()[contains(.,'" + last_4 + "')]]").click()
        browser.find_element_by_id("ctl00_BlueContentPanel_ButtonBar1_ButtonBarButton2").click()
    except:
        pass
def timeStamp(inputText):
    return (inputText + " - " + datetime.datetime.now().strftime("%Y/%m/%d") + " - " + datetime.datetime.now().strftime("%I:%M:%S %p"))
def add_funding_account(proxy_id, last4, address_1, address_2, city, state, zipcode):
    # GoTo funding accounts tab
    try:
        browser.find_element_by_id(
            'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
    except:
        pass
    try:
        browser.find_element_by_link_text('ADD NEW FUNDING ACCOUNT').click()
        browser.find_element_by_id(
            'ctl00_BlueContentPanel_PaymentControl1_PaymentTypeRadioButtonList_0').click()
    except:
        pass
    try:
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located(
                (By.ID, "ctl00_BlueContentPanel_PaymentControl1_ProxyIdTextBox_ProxyIdTextBox_InputTextBox"))
        )
        clearSendkey(browser,
                    'ctl00_BlueContentPanel_PaymentControl1_ProxyIdTextBox_ProxyIdTextBox_InputTextBox',
                    proxy_id)
        clearSendkey(browser,
                    'ctl00_BlueContentPanel_PaymentControl1_Last4OfFundingNumberTextBox_Last4OfFundingNumberTextBox_InputTextBox',
                    last4)
        clearSendkey(browser,
                    'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_AddressLine1TextBox_AddressLine1TextBox_InputTextBox',
                    address_1)
        clearSendkey(browser,
                    'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_AddressLine2TextBox_AddressLine2TextBox_InputTextBox',
                    address_2)
        clearSendkey(browser,
                    'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_CityTextBox_CityTextBox_InputTextBox',
                    city)
        browser.find_element_by_id(
            'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_StateDropDownList_DropDownList').send_keys(
            state)
        clearSendkey(browser,
                    'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_PostalCodeTextBox_PostalCodeTextBox_InputTextBox',
                    zipcode)
        browser.find_element_by_id('ctl00_BlueContentPanel_ButtonBar1_ButtonBarButton2').click()
    except:
        pass
    # GoTo funding accounts tab
    try:
        browser.find_element_by_id(
            'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
    except:
        pass
def High_Roller_Counter():
    if ctr == 1000 or ctr == 2000 or ctr == 3000 or ctr == 4000 or ctr == 5000 or ctr == 6000 or ctr == 7000:
        print(timeStamp(str(ctr)))
    else:
        pass
def newCardholdersearch():
    # navi to chardholder search screen and adj card program and prepaid card number
    # Goto Cardholder Search
    browser.get("https://www.visadpsprepaid.com/PAS/Cardholder/CHSearch.aspx?nav=1")
    try:
        # initial instance timer for proper search setup
        if ctr == 1:
            # adjust searchBy dropdown to ProxyID
            searchBy_dropDown = Select(browser.find_element_by_id("ctl00_MainContentPanel_CHSearchCriteriaControl1_SearchByDropDown_DropDownList"))
            searchBy_dropDown.select_by_index(6)
            # adjust card program dropdown to Aline
            #cardProgram_dropDown = Select(browser.find_element_by_id("ctl00_MainContentPanel_CHSearchCriteriaControl1_CardProgramDropDown_DropDownList"))
            #cardProgram_dropDown.select_by_index(0)
            time.sleep(2)
        else:
            # adjust card program dropdown to All
            cardProgram_dropDown = Select(browser.find_element_by_id("ctl00_MainContentPanel_CHSearchCriteriaControl1_CardProgramDropDown_DropDownList"))
            cardProgram_dropDown.select_by_index(0)
            time.sleep(2)
        # input proxy_ID
        clearSendkey(browser,
                     'ctl00_MainContentPanel_CHSearchCriteriaControl1_Value1TextBox_Value1TextBox_InputTextBox',
                     Card_Holder_ProxyID)
        # select search
        browser.find_element_by_id('ctl00_MainContentPanel_CHSearchCriteriaControl1_SearchLinkButton').click()
        # select CH
        WebDriverWait(browser, 5).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="0_rowInd"]/td[3]/a'))
        )
        browser.find_element_by_xpath('//*[@id="0_rowInd"]/td[3]/a').click()

        #cardholder_eID_adj(CH_EmployeeID, "")

        # GoTo funding accounts tab
        try:
            browser.find_element_by_id(
                'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_10_cell').click()
        except:
            pass
        # input "Notes" field
        try:
            workingCH.append(Card_Holder_ProxyID)
            indexer.append("1")
            clearSendkey(browser,
                'ctl00_BlueContentPanel_CHNotesControl1_NotesTextbox_NotesTextbox_InputTextBox',
                CH_NOTE)
            # input "special instructions" field
            # clearSendkey(browser,
            #     'ctl00_BlueContentPanel_CHNotesControl1_SpecialInstructionsTextbox_SpecialInstructionsTextbox_InputTextBox',
            #     employee_ID)
            # select update
            #browser.find_element_by_xpath("//*[text()[contains(.,'SUBMIT')]]").click()
        except:
            pass
    except:
        # appended info if location is not found
        indexer.append("0")
        workingCH.append(Card_Holder_ProxyID)
        pass
def cardholder_eID_adj(employee_ID, alternate_ID_1):
    # clears a Cardholders Alternate ID 1 field and input info to employee ID field
    # input "alternate id 1" field
    try:
        workingCH.append(Card_Holder_ProxyID)
        indexer.append("1")
        clearSendkey(browser,
            'ctl00_BlueContentPanel_ProfileControl1_AdditionalDataControl1_AdditionalDataRepeater_ctl01_AdditionalDataTextBox_AdditionalDataTextBox_InputTextBox',
            alternate_ID_1)
        # input "employee" field
        clearSendkey(browser,
            'ctl00_BlueContentPanel_ProfileControl1_DemographicControl1_EmployeeIdTextBox_EmployeeIdTextBox_InputTextBox',
            employee_ID)
        # select update
        browser.find_element_by_xpath("//*[text()[contains(.,'UPDATE')]]").click()
    except:
        pass
def expUpdate_addressUpdate(last4, month, year, address_1, address_2, city, state, zipcode):
    # GoTo funding accounts tab
    try:
        browser.find_element_by_id(
            'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
    except:
        pass
    try:
        browser.find_element_by_xpath("//*[text()[contains(.,'" + last4 + "')]]").click()
        browser.find_element_by_link_text('EDIT').click()
    except:
        pass
    try:
        # Check if element exist
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located(
                (By.ID, "ctl00_BlueContentPanel_PaymentControl1_ExpirationDateTextbox_MonthTextBox"))
        )
        # Month text box - Expiration date change
        clearSendkey(browser,
                     'ctl00_BlueContentPanel_PaymentControl1_ExpirationDateTextbox_MonthTextBox',
                     month)
        # Year text box - Expiration date change
        clearSendkey(browser,
                     'ctl00_BlueContentPanel_PaymentControl1_ExpirationDateTextbox_YearTextBox',
                     year)
        # Address Line 1 - Adjustment
        clearSendkey(browser,
                     'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_AddressLine1TextBox_AddressLine1TextBox_InputTextBox',
                     address_1)
        # Address Line 2 - Clear Field
        clearSendkey(browser,
                     'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_AddressLine2TextBox_AddressLine2TextBox_InputTextBox',
                     address_2)
        # City field - Adjustment
        clearSendkey(browser,
                     'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_CityTextBox_CityTextBox_InputTextBox',
                     city)
        # State/Province field - Adjustment
        browser.find_element_by_id(
            'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_StateDropDownList_DropDownList').send_keys(
            state)
        # Postal Code field - Adjustment
        clearSendkey(browser,
                     'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_PostalCodeTextBox_PostalCodeTextBox_InputTextBox',
                     zipcode)
        # Submit button click
        browser.find_element_by_id('ctl00_BlueContentPanel_ButtonBar1_ButtonBarButton3').click()
    except:
        pass
    # GoTo funding accounts tab
    try:
        browser.find_element_by_id(
            'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
    except:
        pass

def location_loop_exe():
    indexer = []
    subclientid = []
    workingLocations = []
    accountsList0 = []
    expDate0 = []
    primeStat0 = []
    proxyId0 = []
    loginScreen(browser)
    print(timeStamp("First Login Successful"))
    ctr = 0
    for index, row in df.iterrows():
        selectLocation = (row['SUBCLIENTDESCRIPTION'])
        location_ID = (row['SUBCLIENTIDENTIFIER'])
        ctr += 1
        if ctr == 2:
            start = time.time()
        elif ctr == 3:
            end = time.time()
            diff = int(end - start)
            print("ETA: " + str((diff * (len(df))) // 60) + "min " + str((diff * (len(df))) % 60) + "sec")
        newLocationsearch()
        relogin()
        df2 = pd.DataFrame({
                            "Index": indexer,
                            "Subclient ID": subclientid,
                            "Subclient Desc": workingLocations,
                            "Funding Account": accountsList0,
                            "ExpDate": expDate0,
                            "Primary Status": primeStat0,
                            "ProxyID #": proxyId0
                            })
        #writer = ExcelWriter('Output_File.xlsx')
        df2.to_csv('CSV_Output.csv', sep='\t', index=False)
        #df2.to_csv('CSV_Output.csv', sep=',', encoding='utf-8', index=False)
        #df2.to_excel(writer, 'Sheet1', index=False)
        #writer.save()
        pyautogui.press("shift")

    print(timeStamp("\nJob Complete"))
def cardholder_loop_exe():
    indexer = []
    workingCH = []
    loginScreen(browser)
    print(timeStamp("First Login Successful"))
    ctr = 0
    for index, row in df.iterrows():
        Card_Holder_ProxyID = (row['CH_ProxyID'])
        CH_EmployeeID = (row['CH_EmployeeID'])
        ctr += 1
        if ctr == 2:
            start = time.time()
        elif ctr == 3:
            end = time.time()
            diff = int(end - start)
            print("ETA: " + str((diff * (len(df))) // 60) + "min " + str((diff * (len(df))) % 60) + "sec")
        newCardholdersearch()
        relogin()
        df2 = pd.DataFrame({
                            "Index": indexer,
                            "CH Proxy ID": workingCH,
                            })
        writer = ExcelWriter('Output_File.xlsx')
        df2.to_excel(writer, 'Sheet1', index=False)
        writer.save()

    print(timeStamp("\nJob Complete"))
def cardInv_loop_exe():
    workingLocations = []
    conf_code = []
    loginScreen(browser)
    print(timeStamp("First Login Successful"))
    ctr = 0
    for index, row in df.iterrows():
        selectLocation = (row['Subclient_IDs'])
        ctr += 1
        if ctr == 2:
            start = time.time()
        elif ctr == 3:
            end = time.time()
            diff = int(end - start)
            print("ETA: " + str((diff * (len(df))) // 60) + "min " + str((diff * (len(df))) % 60) + "sec")
        card_Inventory_Order()
        relogin()
        df2 = pd.DataFrame({
                            "Subclient Name": workingLocations,
                            "Confirmation Num": conf_code,
                            })
        #writer = ExcelWriter('Output_File.xlsx')
        df2.to_csv('CSV_Output.csv', sep='\t', index=False)
        #df2.to_csv('CSV_Output.csv', sep=',', encoding='utf-8', index=False)
        #df2.to_excel(writer, 'Sheet1', index=False)
        #writer.save()
        pyautogui.press("shift")

    print(timeStamp("\nJob Complete"))
def chardholder_notes_loop_exe():
    indexer = []
    workingCH = []
    loginScreen(browser)
    print(timeStamp("First Login Successful"))
    ctr = 0
    for index, row in df.iterrows():
        Card_Holder_ProxyID = (row['PRXY'])
        CH_NOTE = (row['NOTE'])
        ctr += 1
        if ctr == 2:
            start = time.time()
        elif ctr == 3:
            end = time.time()
            diff = int(end - start)
            print("ETA: " + str((diff * (len(df))) // 60) + "min " + str((diff * (len(df))) % 60) + "sec")
        newCardholdersearch()
        relogin()
        df2 = pd.DataFrame({
                            "Index": indexer,
                            "CH Proxy ID": workingCH,
                            })
        df2.to_csv('CSV_Output.csv', sep=',', index=False)
        pyautogui.press("numlock")
        # 'numlock'
    print(timeStamp("\nJob Complete"))
def locationFAanalysis_exe():
    indexer = []
    subclientid = []
    workingLocations = []
    accountsList0 = []
    expDate0 = []
    primeStat0 = []
    proxyId0 = []
    loginScreen(browser)
    print(timeStamp("First Login Successful"))
    ctr = 0
    for index, row in df.iterrows():
        selectLocation = (row['SUBCLIENTDESCRIPTION'])
        location_ID = (row['SUBCLIENTIDENTIFIER'])
        ctr += 1
        if ctr == 2:
            start = time.time()
        elif ctr == 3:
            end = time.time()
            diff = int(end - start)
            print("ETA: " + str((diff * (len(df))) // 60) + "min " + str((diff * (len(df))) % 60) + "sec")
        newLocationsearch()
        relogin()
        df2 = pd.DataFrame({
                            "Index": indexer,
                            "Subclient ID": subclientid,
                            "Subclient Desc": workingLocations,
                            "Funding Account": accountsList0,
                            "ExpDate": expDate0,
                            "Primary Status": primeStat0,
                            "ProxyID #": proxyId0
                            })
        df2.to_csv('CSV_Output.csv', sep=',', index=False)
        pyautogui.press("numlock")
        # 'numlock'
    print(timeStamp("\nJob Complete"))

indexer = []
subclientid = []
workingLocations = []
accountsList0 = []
expDate0 = []
primeStat0 = []
proxyId0 = []
loginScreen(browser)
print(timeStamp("First Login Successful"))
ctr = 0
for index, row in df.iterrows():
    selectLocation = (row['SUBCLIENTDESCRIPTION'])
    location_ID = (row['SUBCLIENTIDENTIFIER'])
    ctr += 1
    if ctr == 2:
        start = time.time()
    elif ctr == 3:
        end = time.time()
        diff = int(end - start)
        print("ETA: " + str((diff * (len(df))) // 60) + "min " + str((diff * (len(df))) % 60) + "sec")
    newLocationsearch()
    relogin()
    df2 = pd.DataFrame({
                        "Index": indexer,
                        "Subclient ID": subclientid,
                        "Subclient Desc": workingLocations,
                        "Funding Account": accountsList0,
                        "ExpDate": expDate0,
                        "Primary Status": primeStat0,
                        "ProxyID #": proxyId0
                        })
    df2.to_csv('CSV_Output.csv', sep=',', index=False)
    pyautogui.press("numlock")
    # 'numlock'
print(timeStamp("\nJob Complete"))
