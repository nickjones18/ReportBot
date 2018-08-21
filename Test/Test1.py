from selenium import webdriver
import time
import getpass
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options

accountOptions = []
campaignOptions = []
touchpointOptions = []

options = Options()
#options.add_argument('--headless')
#options.add_argument('--disable-gpu')
driver = webdriver.Chrome(r"C:\Users\njones\PycharmProjects\ReportBot\Drivers\chromedriver.exe", chrome_options=options)

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
    time.sleep(0.1)
    print(f"{n}. {campaignOptions[i]}")
    n += 1
    i += 1

campaignNumber = input("Select by number: ")

campaign = str(campaignOptions[int(campaignNumber)-1])

# Select Campaign
driver.find_element_by_xpath(f"//select[@id='ctl00_ApplicationContent_MarketingCampaignDropDown']/option[text()='{campaign}']").click()

print("Please wait. This could take a couple minutes.")

# Select Performance
driver.find_element_by_xpath('//*[@title="Blank Email Performance Report" and text() = "Email Performance"]').click()

# Edit
driver.find_element_by_id("ctl00_ApplicationContent_hlShowDetails").click()

#Prompt for touchpoint
select = Select(driver.find_element_by_id("ctl00_ApplicationContent_ReportID50_TouchpointDropDown"))
opts = select.options
for opt in opts:
    touchpointOptions.append(opt.text)

print("Please choose a touchpoint from the following list...")
n = 1
for i in range (len(touchpointOptions)):
    time.sleep(0.1)
    print(f"{n}. {touchpointOptions[i]}")
    n += 1
    i += 1

touchpointNumber = input("Select by number: ")

touchpoint = str(touchpointOptions[int(touchpointNumber)-1])


driver.find_element_by_xpath(f"//select[@id='ctl00_ApplicationContent_ReportID50_TouchpointDropDown']/option[text()='{touchpoint}']").click()

print("Pulling the reports from the last deployment! They should be in your downloads folder momentarily.")

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

