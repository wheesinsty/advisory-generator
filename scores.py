import playwright
from playwright.sync_api import sync_playwright
import pandas as pd
import os

# opens the excel sheets to obtain the name and the scores
def openExcel():
    return pd.read_excel("C:/Users/" + os.getenv('USERNAME') + "/Documents/AdvisoryAutomate/test.xlsx", sheet_name = "Report status"), pd.read_excel("C:/Users/" + os.getenv('USERNAME') + "/Documents/AdvisoryAutomate/test.xlsx", sheet_name = "Adoption score", index_col = 0)

# reports any errors
def reportError(msg):
    # report error to excel sheet
    scoreSheet.loc[client, "Error"] = scoreSheet.loc[client, "Error"] + "\n " + msg
    with pd.ExcelWriter("test.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        scoreSheet.to_excel(writer, sheet_name='Adoption score', index=False)

# goes to the client's admin centre
def goToTenant(companiesSheet, client) -> bool:
    
    # goes to the customer list
    page.goto("https://partner.microsoft.com/dashboard/v2/customers/list")
        
    # fill customer name into searchbar
    page.locator("#customer-search-box").get_by_placeholder("Search").fill(companiesSheet.loc[client]["Domain"])
    page.wait_for_timeout(5000)
    
    # go to customer profile
    page.get_by_role("radio", name = "Select row", checked = False, disabled = False).check()
    
    # go to service management
    page.get_by_text("Service management").click()
    page.wait_for_timeout(5000)
    
    # go to "Microsoft 365"
    try:
        page.locator('//*[@id="MicrosoftOffice"]').click()
        page.wait_for_timeout(5000)
    except:
        reportError("No admin permissions")
        return True

    return False

# extracts scores from the adoption score page
def getScores(company, scores):
    page.goto("https://admin.microsoft.com/Adminportal/Home#/adoptionscore")
    
    # extracts "Your organization's score"
    scoreSheet.loc[company, 'Your organization’s score'] = page.get_by_text("Your organization’s score: ").text_content()[-3:]
    
    # extracts "Total score"
    scoreSheet.loc[company, 'Total score'] = page.get_by_text("Total score:").all_text_contents()[0][-14:-7]
    
    # extracts everything else
    for link in range(len(scores)):
        if link < 6:
            try:
                scoreSheet.loc[company, scores[link]] = page.locator('//html/body/div[1]/div[1]/div[1]/main/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div[2]/div[' + str(link+1) + ']/div/div[1]/span[3]').text_content().strip(" points")
            except:
                scoreSheet.loc[company, scores[link]] = '--'
        else:
            try:
                scoreSheet.loc[company, scores[link]] = page.locator('//html/body/div[1]/div[1]/div[1]/main/div[2]/div[1]/div[2]/div[2]/div/div/div[2]/div[2]/div[' + str(link-5) + ']/div/div[1]/span[3]').text_content().strip(" points")
            except:
                scoreSheet.loc[company, scores[link]] = '--'
    print(scoreSheet)
    with pd.ExcelWriter("test.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        scoreSheet.to_excel(writer, sheet_name='Adoption score', index=False)
                                                           
companiesSheet, scoreSheet = openExcel()

# list of scores to extract
scores = ['Communication', 'Meetings', 'Content collaboration', 'Teamwork', 'Mobility', 'AI assistance', 'Endpoint analytics',	'Network connectivity', 'Microsoft 365 Apps Health']

for client in range(len(companiesSheet["Domain"])): # client is a number
    
    company = companiesSheet.loc[client, "Username"]
    
    # adds a new row to the df
    scoreSheet.loc[company] = [None] * len(scoreSheet.columns)

    with sync_playwright() as p:
        # connects to open browser
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        default_context = browser.contexts[0]
        page = default_context.pages[0]
        
        # goes to the client's admin centre
        if goToTenant(companiesSheet, client): continue
        
        # extracts the scores
        getScores(company, scores)
    
with pd.ExcelWriter("test.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    scoreSheet.to_excel(writer, sheet_name='Adoption score', index=False)
