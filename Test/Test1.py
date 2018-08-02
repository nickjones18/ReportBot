from selenium import webdriver
import time
from selenium.webdriver.support.ui import Select

accountOptions = []
campaignOptions = []

driver = webdriver.Chrome(r"C:\Users\njones\PycharmProjects\ReportBot\Drivers\chromedriver.exe")
# driver.set_page_load_timeout(10)
driver.get("http://uweb01/MarketingConsole/Login.aspx?ReturnUrl=%2fMarketingConsole%2fDefault.aspx")

#Get the user's input
username = input("Username: ")
password = input("Password: ")
print("Logging in...")
# Login
driver.find_element_by_id("AccountLogin_Login1_UserName").send_keys(f"{username}")
driver.find_element_by_id("AccountLogin_Login1_Password").send_keys(f"{password}")
driver.find_element_by_id("AccountLogin_Login1_LoginButton").click()

#Prompt for account
select = Select(driver.find_element_by_id("ctl00_ApplicationContent_AccountDropDown"))
opts = select.options
for opt in opts:
    accountOptions.append(opt.text)

print("Please choose an account from the following list...")
n = 1
for i in range (len(accountOptions)):
    time.sleep(0.1)
    print(f"{n}. {accountOptions[i]}")
    n += 1
    i += 1

accountNumber = input("Select by number: ")

account = str(accountOptions[int(accountNumber)-1])

# Select Account
driver.find_element_by_xpath(f"//select[@id='ctl00_ApplicationContent_AccountDropDown']/option[text()='{account}']").click()

#Prompt for campaign
select = Select(driver.find_element_by_id("ctl00_ApplicationContent_MarketingCampaignDropDown"))
opts = select.options
for opt in opts:
    campaignOptions.append(opt.text)

print("Please choose a campaign from the following list...")
n = 1
for i in range (len(campaignOptions)):
    time.sleep(0.5)
    print(f"{n}. {campaignOptions[i]}")
    n += 1
    i += 1

campaignNumber = input("Select by number: ")

campaign = str(accountOptions[int(accountNumber)-1])

touchpoint = input("Touchpoint: ")

# Select Campaign
driver.find_element_by_xpath(f"//select[@id='ctl00_ApplicationContent_MarketingCampaignDropDown']/option[text()='{campaign}']").click()

# Select Performance
driver.find_element_by_xpath('//*[@title="Blank Email Performance Report" and text() = "Email Performance"]').click()

# Edit
driver.find_element_by_id("ctl00_ApplicationContent_hlShowDetails").click()

# Edit Touchpoint
driver.find_element_by_xpath(f"//select[@id='ctl00_ApplicationContent_ReportID50_TouchpointDropDown']/option[text()='{touchpoint}']").click()

# Download Performance
driver.find_element_by_id("ctl00_ApplicationContent_ExportToPNGButton").click()

# Select Email Failure Population List
driver.find_element_by_xpath('//*[@title="Lists the recipients whose email failed to be delivered." and text() = "Email Failure Population List"]').click()

# Edit
driver.find_element_by_id("ctl00_ApplicationContent_hlShowDetails").click()

# Edit Touchpoint
driver.find_element_by_id("ctl00_ApplicationContent_ReportID7_FilterPanelSeriesA_ConditionPanel1_DynamicControl96be917e_e47b_4356_b0b4_03424a3b5777_EmailEventFilter1_TouchPointList_Arrow").click()
driver.find_element_by_xpath(f"//*[text()='{touchpoint}']").click()

# Download Failure List
driver.find_element_by_id("ctl00_ApplicationContent_ExportToExcelButton").click()

# Select Email Clicks List
driver.find_element_by_xpath('//*[@title="Lists the recipients who clicked a link in an email." and text() = "Email Clicks Population List"]').click()

# Edit
driver.find_element_by_id("ctl00_ApplicationContent_hlShowDetails").click()

# Edit Touchpoint
driver.find_element_by_id("ctl00_ApplicationContent_ReportID7_FilterPanelSeriesA_ConditionPanel1_DynamicControl96be917e_e47b_4356_b0b4_03424a3b5777_EmailEventFilter1_TouchPointList_Arrow").click()
driver.find_element_by_xpath(f"//*[text()='{touchpoint}']").click()

# Download Failure List
driver.find_element_by_id("ctl00_ApplicationContent_ExportToExcelButton").click()

# Select Email Unsubscribe List
driver.find_element_by_xpath('//*[@title="Lists the recipients who unsubscribed from email notifications." and text() = "Email Unsubscribed Population List"]').click()

# Edit
driver.find_element_by_id("ctl00_ApplicationContent_hlShowDetails").click()

# Edit Touchpoint
driver.find_element_by_id("ctl00_ApplicationContent_ReportID7_FilterPanelSeriesA_ConditionPanel1_DynamicControl96be917e_e47b_4356_b0b4_03424a3b5777_EmailEventFilter1_TouchPointList_Arrow").click()
driver.find_element_by_xpath(f"//*[text()='{touchpoint}']").click()

# Download Unsubscribe List
driver.find_element_by_id("ctl00_ApplicationContent_ExportToExcelButton").click()

# Close Driver
# driver.quit()

