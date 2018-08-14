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

class VisaPAS:
    indexer = []
    subclientid = []
    workingLocations = []
    accountsList0 = []
    expDate0 = []
    primeStat0 = []
    proxyId0 = []
    conf_code = []
    checker = []
    card_programs = {"ADP Aline": 1,
                    "ADP Aline AP": 2,
                    "ADP Aline-MC": 3,
                    "ADP TotalPay Card": 4,
                    "ADP TotalPay Card GPR-FDIC": 5,
                    "ADP TotalPay Conv": 6,
                    "Aline": 7,
                    "Aline-CNB": 8,
                    "Aline-MB": 9,
                    "Aline-MC": 10,
                    "Wisely": 11}
    browser = webdriver.Chrome(executable_path="C:\\Users\\reyesa1\\Documents\\chromedriver.exe")
    df = pd.read_excel(open("C:\\Users\\reyesa1\\Documents\\Atom Projects\\Input_data.xlsx",'rb'), sheet_name=0)
    def __init__(self, subclient_desc, location_ID):
        self.subclient_desc = subclient_desc
        self.location_ID = location_ID
        # self.invAmount = invAmount
        # self.invNote = invNote
    def login(self):
        VisaPAS.browser.get("https://www.visadpsprepaid.com/pas")
        visa_id = input("Enter username: prc562.")
        visa_pwd = input("Enter password: ")
        username = VisaPAS.browser.find_element_by_name('USER')
        username.send_keys("prc562." + visa_id)
        password = VisaPAS.browser.find_element_by_name('PASSWORD')
        password.send_keys(visa_pwd)
        login = VisaPAS.browser.find_element_by_id('SubmitButton').click()
        print("First Login Successful" + " - " + datetime.datetime.now().strftime("%Y/%m/%d") + " - " + datetime.datetime.now().strftime("%I:%M:%S %p"))
    def relogin(self):
        try:
            VisaPAS.browser.find_element_by_id('PleaseLogOn')
            VisaPAS.browser.get("https://www.visadpsprepaid.com/pas")
            username = VisaPAS.browser.find_element_by_name('USER')
            username.send_keys("prc562." + visa_id)
            password = VisaPAS.browser.find_element_by_name('PASSWORD')
            password.send_keys(visa_pwd)
            login = VisaPAS.browser.find_element_by_id('SubmitButton').click()
            print("*****Relogin Initiated*****" + " - " + datetime.datetime.now().strftime("%Y/%m/%d") + " - " + datetime.datetime.now().strftime("%I:%M:%S %p"))
        except:
            pass
    def location_search(self):
        VisaPAS.browser.get("https://www.visadpsprepaid.com/PAS/Location/LOCSearch.aspx?nav=1")
        #print(self.subclient_desc + "Andrew" + self.location_ID)
        try:
            clearSendkey(VisaPAS.browser,
                         'ctl00_MainContentPanel_LOCSearchCriteriaControl1_LocationNameTextBox_LocationNameTextBox_InputTextBox',
                         self.subclient_desc)
            VisaPAS.browser.find_element_by_id('ctl00_MainContentPanel_LOCSearchCriteriaControl1_SearchLinkButton').click()
        except:
            pass
        try:
            # locate location and select after search
            try:
                VisaPAS.browser.find_element_by_xpath('//*[@id="0_rowInd"]/td[1]/a').click()
                sub_id = VisaPAS.browser.find_element_by_id('ctl00_BlueContentPanel_LOCTypeControl1_lblLocationIdValue').text
                if ''.join(sub_id.split()) == ''.join(self.location_ID.split()):
                    pass
                else:
                    VisaPAS.browser.get("https://www.visadpsprepaid.com/PAS/Location/LOCSearch.aspx?nav=1")
                    VisaPAS.browser.find_element_by_xpath('//*[@id="1_rowInd"]/td[1]/a').click()
                    sub_id = VisaPAS.browser.find_element_by_id('ctl00_BlueContentPanel_LOCTypeControl1_lblLocationIdValue').text
                    if ''.join(sub_id.split()) == ''.join(self.location_ID.split()):
                        pass
                    else:
                        VisaPAS.browser.get("https://www.visadpsprepaid.com/PAS/Location/LOCSearch.aspx?nav=1")
                        VisaPAS.browser.find_element_by_xpath('//*[@id="2_rowInd"]/td[1]/a').click()
                        sub_id = VisaPAS.browser.find_element_by_id('ctl00_BlueContentPanel_LOCTypeControl1_lblLocationIdValue').text
                        if ''.join(sub_id.split()) == ''.join(self.location_ID.split()):
                            pass
                        else:
                            pass
            except:
                pass
        except:
            pass
    def Funding_Account_Analysis(self):
        # GoTo funding accounts tab
        try:
            VisaPAS.browser.find_element_by_id(
            'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
        except:
            pass
        # data collection code of funding accounts
        invalid_search = "There are no items to show in this view."
        try:
            invalid_check = VisaPAS.browser.find_element_by_xpath(
                '//*[@id="ctl00_MainContentPanel_LOCSearchResultsControl1_grdResults"]/tbody/tr[2]/td/span').get_attribute(
                'textContent')
        except:
            invalid_check = "None"
            pass
        if invalid_check == invalid_search:
            VisaPAS.indexer.append("0")
            VisaPAS.subclientid.append(self.location_ID)
            VisaPAS.workingLocations.append(self.subclient_desc)
            VisaPAS.accountsList0.append("N/A")
            VisaPAS.expDate0.append("N/A")
            VisaPAS.primeStat0.append("N/A")
            VisaPAS.proxyId0.append("N/A")
        else:
            try:
                VisaPAS.subclientid.append(self.location_ID)
                VisaPAS.workingLocations.append(self.subclient_desc)
                VisaPAS.indexer.append("1")
                VisaPAS.accountsList0.append(VisaPAS.browser.find_element_by_id(
                    'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl02_lblAccountNumber').text)
                VisaPAS.expDate0.append(VisaPAS.browser.find_element_by_id(
                    'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl02_lblExpDate').text)
                VisaPAS.primeStat0.append(VisaPAS.browser.find_element_by_id(
                    'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl02_lblPrimaryAccount').text)
                VisaPAS.proxyId0.append(VisaPAS.browser.find_element_by_xpath(
                    '//*[@id="ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl02_ctl06_tblInfo"]/tbody/tr[1]/td[2]').get_attribute(
                    'textContent').strip())
            except:
                VisaPAS.accountsList0.append("N/A")
                VisaPAS.expDate0.append("N/A")
                VisaPAS.primeStat0.append("N/A")
                VisaPAS.proxyId0.append("N/A")
                pass
            try:
                VisaPAS.accountsList0.append(VisaPAS.browser.find_element_by_id(
                    'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl03_lblAccountNumber').text)
                VisaPAS.expDate0.append(VisaPAS.browser.find_element_by_id(
                    'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl03_lblExpDate').text)
                VisaPAS.primeStat0.append(VisaPAS.browser.find_element_by_id(
                    'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl03_lblPrimaryAccount').text)
                VisaPAS.proxyId0.append(VisaPAS.browser.find_element_by_xpath(
                    '//*[@id="ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl03_ctl06_tblInfo"]/tbody/tr[1]/td[2]').get_attribute(
                    'textContent').strip())
                VisaPAS.subclientid.append(self.location_ID)
                VisaPAS.workingLocations.append(self.subclient_desc)
                VisaPAS.indexer.append("2")
            except:
                pass
            try:
                VisaPAS.accountsList0.append(VisaPAS.browser.find_element_by_id(
                    'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl04_lblAccountNumber').text)
                VisaPAS.expDate0.append(VisaPAS.browser.find_element_by_id(
                    'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl04_lblExpDate').text)
                VisaPAS.primeStat0.append(VisaPAS.browser.find_element_by_id(
                    'ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl04_lblPrimaryAccount').text)
                VisaPAS.proxyId0.append(VisaPAS.browser.find_element_by_xpath(
                    '//*[@id="ctl00_BlueContentPanel_LOCFundingAccountsListControl1_grdResults_ctl04_ctl06_tblInfo"]/tbody/tr[1]/td[2]').get_attribute(
                    'textContent').strip())
                VisaPAS.subclientid.append(self.location_ID)
                VisaPAS.workingLocations.append(self.subclient_desc)
                VisaPAS.indexer.append("3")
            except:
                pass
        df2 = pd.DataFrame({
                            "Index": VisaPAS.indexer,
                            "Subclient ID": VisaPAS.subclientid,
                            "Subclient Desc": VisaPAS.workingLocations,
                            "Funding Account": VisaPAS.accountsList0,
                            "ExpDate": VisaPAS.expDate0,
                            "Primary Status": VisaPAS.primeStat0,
                            "ProxyID #": VisaPAS.proxyId0
                            })
        df2.to_csv('CSV_Output.csv', sep=',', index=False)
        pyautogui.press("numlock")
    def expiration_address_update(self, last4, month, year, address_1, address_2, city, state, zipcode):
        # GoTo funding accounts tab
        try:
            VisaPAS.browser.find_element_by_id(
                'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
        except:
            pass
        try:
            VisaPAS.browser.find_element_by_xpath("//*[text()[contains(.,'" + last4 + "')]]").click()
            VisaPAS.browser.find_element_by_link_text('EDIT').click()
        except:
            pass
        try:
            # Check if element exist
            WebDriverWait(VisaPAS.browser, 10).until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_BlueContentPanel_PaymentControl1_ExpirationDateTextbox_MonthTextBox"))
            )
            # Month text box - Expiration date change
            clearSendkey(VisaPAS.browser,
                         'ctl00_BlueContentPanel_PaymentControl1_ExpirationDateTextbox_MonthTextBox',
                         month)
            # Year text box - Expiration date change
            clearSendkey(VisaPAS.browser,
                         'ctl00_BlueContentPanel_PaymentControl1_ExpirationDateTextbox_YearTextBox',
                         year)
            # Address Line 1 - Adjustment
            clearSendkey(VisaPAS.browser,
                         'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_AddressLine1TextBox_AddressLine1TextBox_InputTextBox',
                         address_1)
            # Address Line 2 - Clear Field
            clearSendkey(VisaPAS.browser,
                         'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_AddressLine2TextBox_AddressLine2TextBox_InputTextBox',
                         address_2)
            # City field - Adjustment
            clearSendkey(VisaPAS.browser,
                         'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_CityTextBox_CityTextBox_InputTextBox',
                         city)
            # State/Province field - Adjustment
            VisaPAS.browser.find_element_by_id(
                'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_StateDropDownList_DropDownList').send_keys(
                state)
            # Postal Code field - Adjustment
            clearSendkey(VisaPAS.browser,
                         'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_PostalCodeTextBox_PostalCodeTextBox_InputTextBox',
                         zipcode)
            # Submit button click
            VisaPAS.browser.find_element_by_id('ctl00_BlueContentPanel_ButtonBar1_ButtonBarButton3').click()
        except:
            pass
        # GoTo funding accounts tab
        try:
            VisaPAS.browser.find_element_by_id(
                'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
        except:
            pass
    def set_FA_primary(self, last_4):
        # GoTo funding accounts tab
        try:
            VisaPAS.browser.find_element_by_id(
            'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
        except:
            pass
        try:
            VisaPAS.browser.find_element_by_xpath("//*[text()[contains(.,'" + last_4 + "')]]").click()
            VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_ButtonBar1_ButtonBarButton2").click()
        except:
            pass
    def add_funding_account(self, proxy_id, last4, address_1, address_2, city, state, zipcode):
        # GoTo funding accounts tab
        try:
            VisaPAS.browser.find_element_by_id(
                'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
        except:
            pass
        try:
            VisaPAS.browser.find_element_by_link_text('ADD NEW FUNDING ACCOUNT').click()
            VisaPAS.browser.find_element_by_id(
                'ctl00_BlueContentPanel_PaymentControl1_PaymentTypeRadioButtonList_0').click()
        except:
            pass
        try:
            WebDriverWait(VisaPAS.browser, 10).until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_BlueContentPanel_PaymentControl1_ProxyIdTextBox_ProxyIdTextBox_InputTextBox"))
            )
            clearSendkey(VisaPAS.browser,
                        'ctl00_BlueContentPanel_PaymentControl1_ProxyIdTextBox_ProxyIdTextBox_InputTextBox',
                        proxy_id)
            clearSendkey(VisaPAS.browser,
                        'ctl00_BlueContentPanel_PaymentControl1_Last4OfFundingNumberTextBox_Last4OfFundingNumberTextBox_InputTextBox',
                        last4)
            clearSendkey(VisaPAS.browser,
                        'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_AddressLine1TextBox_AddressLine1TextBox_InputTextBox',
                        address_1)
            clearSendkey(VisaPAS.browser,
                        'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_AddressLine2TextBox_AddressLine2TextBox_InputTextBox',
                        address_2)
            clearSendkey(VisaPAS.browser,
                        'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_CityTextBox_CityTextBox_InputTextBox',
                        city)
            VisaPAS.browser.find_element_by_id(
                'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_StateDropDownList_DropDownList').send_keys(
                state)
            clearSendkey(VisaPAS.browser,
                        'ctl00_BlueContentPanel_PaymentControl1_AddressControl1_PostalCodeTextBox_PostalCodeTextBox_InputTextBox',
                        zipcode)
            VisaPAS.browser.find_element_by_id('ctl00_BlueContentPanel_ButtonBar1_ButtonBarButton2').click()
        except:
            pass
        # GoTo funding accounts tab
        try:
            VisaPAS.browser.find_element_by_id(
                'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
        except:
            pass
    def remove_funding_Account(self, account_number):
        # GoTo funding accounts tab
        try:
            VisaPAS.browser.find_element_by_id(
                'ctl00_SideNavPanel_SideNavMenuControl_LeftNavigationPortalContextMenu_LeftNavigationNavBar_item_4_cell').click()
        except:
            pass
        try:
            VisaPAS.browser.find_element_by_xpath("//*[text()[contains(.,'" + account_number + "')]]").click()
            VisaPAS.browser.find_element_by_xpath("//*[text()[contains(.,'EDIT')]]").click()
            VisaPAS.browser.find_element_by_xpath("//*[text()[contains(.,'DELETE ACCOUNT')]]").click()
        except:
            pass
        try:
            WebDriverWait(VisaPAS.browser, 3).until(EC.alert_is_present(),
                                            'Timed out waiting for PA creation ' +
                                            'confirmation popup to appear.')
            alert = VisaPAS.browser.switch_to.alert
            alert.accept()
        except:
            pass
    def card_Inventory_Order(self, card_program, invAmount, invNote):
        VisaPAS.browser.get("https://www.visadpsprepaid.com/PAS/InventoryControl/ICCreateOrder.aspx")
        try:
            VisaPAS.workingLocations.append(self.subclient_desc)
            VisaPAS.subclientid.append(self.location_ID)
            # create a dictionary for the card programs {Aline-MB: 9}...
            # adjust card program dropdown to Aline-MB
            WebDriverWait(VisaPAS.browser, 20).until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_BlueContentPanel_ICCardOrderHeader1_ICCardSelectorControl1_DropDownCardProgram_DropDownList"))
            )
            card_programDropdown = Select(VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_ICCardOrderHeader1_ICCardSelectorControl1_DropDownCardProgram_DropDownList"))
            card_programDropdown.select_by_index(VisaPAS.card_programs[card_program])

            clearSendkey(VisaPAS.browser,
                         'ctl00_BlueContentPanel_ICCardOrderHeader1_ICCardSelectorControl1_DropDownLocation_SearchTextBox_SearchTextBox_InputTextBox',
                         self.subclient_desc)
            time.sleep(2)
            location_select_Dropdown = Select(VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_ICCardOrderHeader1_ICCardSelectorControl1_DropDownLocation_DropDownList"))
            location_select_Dropdown.select_by_index(1)
        except:
            pass
        try:
            WebDriverWait(VisaPAS.browser, 20).until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_BlueContentPanel_ICCardOrderInformation1_DropDownCardDesign_DropDownList"))
            )
            card_design_Dropdown = Select(VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_ICCardOrderInformation1_DropDownCardDesign_DropDownList"))
            card_design_Dropdown.select_by_index(1)

            VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_ICCardOrderInformation1_TextBoxCardQuantity_TextBoxCardQuantity_InputTextBox").send_keys(invAmount)
            VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_ICCardOrderInformation1_TextBoxComments").send_keys(invNote)
            VisaPAS.browser.find_element_by_xpath("//*[text()[contains(.,'SUBMIT')]]").click()
            try:
                WebDriverWait(VisaPAS.browser, 3).until(EC.alert_is_present(),
                                                'Timed out waiting for PA creation ' +
                                                'confirmation popup to appear.')
                alert = VisaPAS.browser.switch_to.alert
                alert.accept()
            except:
                pass
            VisaPAS.conf_code.append(VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_ICOrderQueueSearchCriteria1_TextBoxOrderNumber_TextBoxOrderNumber_InputTextBox").get_attribute("value"))
        except:
            VisaPAS.conf_code.append("error")
            pass
        df2 = pd.DataFrame({
                            "Subclient ID": VisaPAS.subclientid,
                            "Subclient Name": VisaPAS.workingLocations,
                            "Confirmation Num": VisaPAS.conf_code,
                            })
        df2.to_csv('CSV_Output.csv', sep=',', index=False)
        pyautogui.press("numlock")
    def card_inv_sub_adjust(self, card_programs):
        """This module goes into the location settings and modifies the  """
        VisaPAS.browser.get("https://www.visadpsprepaid.com/PAS/InventoryControl/ICLocationSettingsManager.aspx")
        try:
            VisaPAS.workingLocations.append(self.subclient_desc)
            VisaPAS.subclientid.append(self.location_ID)
            # create a dictionary for the card programs {Aline-MB: 9}...
            # adjust card program dropdown to Aline-MB
            WebDriverWait(VisaPAS.browser, 20).until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_BlueContentPanel_IClocationCardStockHeader1_ICCardSelectorControl1_DropDownCardProgram_DropDownList"))
            )
            card_programDropdown = Select(VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_IClocationCardStockHeader1_ICCardSelectorControl1_DropDownCardProgram_DropDownList"))
            card_programDropdown.select_by_index(card_programs)

            clearSendkey(VisaPAS.browser,
                         'ctl00_BlueContentPanel_IClocationCardStockHeader1_ICCardSelectorControl1_DropDownLocation_SearchTextBox_SearchTextBox_InputTextBox',
                         self.subclient_desc)
            time.sleep(2)
            location_select_Dropdown = Select(VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_IClocationCardStockHeader1_ICCardSelectorControl1_DropDownLocation_DropDownList"))
            location_select_Dropdown.select_by_index(1)
            VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_IClocationCardStockHeader1_ButtonBar1_DefaultButton1").click()
        except:
            pass
        try:
            WebDriverWait(VisaPAS.browser, 20).until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_BlueContentPanel_ICLocationCardStock1_ButtonBar1_DefaultButton1"))
            )
            sub_account = VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_ICLocationCardStock1_grdResults_ctl02_chkPasCheckField_0")
            if sub_account.is_selected():
                VisaPAS.checker.append("1")
                sub_account.click()
                VisaPAS.browser.find_element_by_xpath("//*[text()[contains(.,'SUBMIT')]]").click()
                try:
                    WebDriverWait(VisaPAS.browser, 3).until(EC.alert_is_present(),
                                                    'Timed out waiting for PA creation ' +
                                                    'confirmation popup to appear.')
                    alert = VisaPAS.browser.switch_to.alert
                    alert.accept()
                except:
                    pass
            else:
                VisaPAS.checker.append("0")
            VisaPAS.conf_code.append("Successful")
        except:
            VisaPAS.conf_code.append("error")
            VisaPAS.checker.append("n/a")
            pass
        df2 = pd.DataFrame({
                            "Subclient ID": VisaPAS.subclientid,
                            "Subclient Name": VisaPAS.workingLocations,
                            "Confirmation Num": VisaPAS.conf_code,
                            "is selected?": VisaPAS.checker
                            })
        df2.to_csv('CSV_Output.csv', sep=',', index=False)
        pyautogui.press("numlock")
    def auto_replenishment(self, card_program):
        """This module goes into the location settings and modifies the  """
        VisaPAS.browser.get("https://www.visadpsprepaid.com/PAS/InventoryControl/ICLocationSettingsManager.aspx")
        try:
            VisaPAS.workingLocations.append(self.subclient_desc)
            VisaPAS.subclientid.append(self.location_ID)
            # create a dictionary for the card programs {Aline-MB: 9}...
            # adjust card program dropdown to Aline-MB
            WebDriverWait(VisaPAS.browser, 20).until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_BlueContentPanel_IClocationCardStockHeader1_ICCardSelectorControl1_DropDownCardProgram_DropDownList"))
            )
            card_programDropdown = Select(VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_IClocationCardStockHeader1_ICCardSelectorControl1_DropDownCardProgram_DropDownList"))
            card_programDropdown.select_by_index(VisaPAS.card_programs[card_program])

            clearSendkey(VisaPAS.browser,
                         'ctl00_BlueContentPanel_IClocationCardStockHeader1_ICCardSelectorControl1_DropDownLocation_SearchTextBox_SearchTextBox_InputTextBox',
                         self.subclient_desc)
            time.sleep(2)
            location_select_Dropdown = Select(VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_IClocationCardStockHeader1_ICCardSelectorControl1_DropDownLocation_DropDownList"))
            location_select_Dropdown.select_by_index(1)
            VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_IClocationCardStockHeader1_ButtonBar1_DefaultButton1").click()
        except:
            pass
        try:
            VisaPAS.browser.find_element_by_xpath("//*[text()[contains(.,'ALINE MC Access English')]]").click()
            VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_ICLocationCardStock1_ButtonBar1_ButtonBarButton1").click()
            WebDriverWait(VisaPAS.browser, 20).until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_BlueContentPanel_ICLocationEditCardStockSettings1_RadioButtonAutoApproveOn"))
            )
            VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_ICLocationEditCardStockSettings1_RadioButtonAutoApproveOn").click()
            VisaPAS.browser.find_element_by_id("ctl00_BlueContentPanel_ButtonBar1_DefaultButton1").click()
        except:
            pass
    def tearDown(self):
        print("Job Complete || closing browser in T-10")
        time.sleep(10)
        VisaPAS.browser.close()


VisaPAS("","").login()
for index, row in VisaPAS.df.iterrows():
    Visa = VisaPAS(row['SUBCLIENTDESCRIPTION'], row['SUBCLIENTIDENTIFIER'])
    Visa.location_search()
    Visa.relogin()
Visa.tearDown()
