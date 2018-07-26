from selenium import webdriver
import time

driver = webdriver.Chrome(r"C:\Users\njones\PycharmProjects\ReportBot\Drivers\chromedriver.exe")

#driver.set_page_load_timeout(10)

driver.get("http://uweb01/MarketingConsole/Login.aspx?ReturnUrl=%2fMarketingConsole%2fDefault.aspx")

# Prompt for username and password
username = input("Please enter your username: ")
password = input("Please enter your password: ")
account = input("Please enter account name: ")
campaign = input("Please enter campaign name: ")
touchpoint = input("Please enter touchpoint: ")

# Login
driver.find_element_by_id(f"AccountLogin_Login1_UserName").send_keys("{username}")
driver.find_element_by_id(f"AccountLogin_Login1_Password").send_keys("{password}")
driver.find_element_by_id("AccountLogin_Login1_LoginButton").click()

#Select Account
driver.find_element_by_xpath(f"//select[@id='ctl00_ApplicationContent_AccountDropDown']/option[text()='{account}']").click()

#Select Account
driver.find_element_by_xpath(f"//select[@id='ctl00_ApplicationContent_MarketingCampaignDropDown']/option[text()='{campaign}']").click()

#Select Performance
driver.find_element_by_xpath('//*[@title="Blank Email Performance Report" and text() = "Email Performance"]').click()

#Edit
driver.find_element_by_id("ctl00_ApplicationContent_hlShowDetails").click()

#Edit Touchpoint
driver.find_element_by_xpath(f"//select[@id='ctl00_ApplicationContent_ReportID50_TouchpointDropDown']/option[text()='{touchpoint}']").click()

#Download Performance
driver.find_element_by_id("ctl00_ApplicationContent_ExportToPNGButton").click()

#Select Email Failure Population List
driver.find_element_by_xpath('//*[@title="Lists the recipients whose email failed to be delivered." and text() = "Email Failure Population List"]').click()

# Edit
driver.find_element_by_id("ctl00_ApplicationContent_hlShowDetails").click()

#Edit Touchpoint
driver.find_element_by_id("ctl00_ApplicationContent_ReportID7_FilterPanelSeriesA_ConditionPanel1_DynamicControl96be917e_e47b_4356_b0b4_03424a3b5777_EmailEventFilter1_TouchPointList_Arrow").click()
driver.find_element_by_xpath(f"//*[text()='{touchpoint}']").click()

#Download Failure List
driver.find_element_by_id("ctl00_ApplicationContent_ExportToExcelButton").click()

#Select Email Clicks List
driver.find_element_by_xpath('//*[@title="Lists the recipients who clicked a link in an email." and text() = "Email Clicks Population List"]').click()

#Edit
driver.find_element_by_id("ctl00_ApplicationContent_hlShowDetails").click()

#Edit Touchpoint
driver.find_element_by_id("ctl00_ApplicationContent_ReportID7_FilterPanelSeriesA_ConditionPanel1_DynamicControl96be917e_e47b_4356_b0b4_03424a3b5777_EmailEventFilter1_TouchPointList_Arrow").click()
driver.find_element_by_xpath(f"//*[text()='{touchpoint}']").click()

#Download Failure List
driver.find_element_by_id("ctl00_ApplicationContent_ExportToExcelButton").click()

#Select Email Unsubscribe List
driver.find_element_by_xpath('//*[@title="Lists the recipients who unsubscribed from email notifications." and text() = "Email Unsubscribed Population List"]').click()

#Edit
driver.find_element_by_id("ctl00_ApplicationContent_hlShowDetails").click()

#Edit Touchpoint
driver.find_element_by_id("ctl00_ApplicationContent_ReportID7_FilterPanelSeriesA_ConditionPanel1_DynamicControl96be917e_e47b_4356_b0b4_03424a3b5777_EmailEventFilter1_TouchPointList_Arrow").click()
driver.find_element_by_xpath(f"//*[text()='{touchpoint}']").click()

#Download Unsubscribe List
driver.find_element_by_id("ctl00_ApplicationContent_ExportToExcelButton").click()

#Close Driver
driver.quit()