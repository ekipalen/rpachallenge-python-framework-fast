from time import sleep
from robocorp import browser, excel, http
from robocorp.tasks import task


@task
def solve_challenge():
    """
    Solve the RPA challenge
    
    Change the headless=True > headless=False to archieve a bit better result.  
    """
    browser.configure(
        browser_engine="chromium",
        screenshot="only-on-failure",
        headless=False,
    )
        
    http.download("https://rpachallenge.com/assets/downloadFiles/challenge.xlsx")
    worksheet = excel.open_workbook("challenge.xlsx").worksheet("Sheet1")

    page = browser.goto("https://rpachallenge.com/")
    page.click("button:text('Start')")

    for row in worksheet.as_table(header=True):
        fill_and_submit_form(row)

    browser.screenshot()
    sleep(10)


def fill_and_submit_form(row):
    page = browser.page()
    page.evaluate(f'''() => {{
        document.evaluate('//input[@ng-reflect-name="labelFirstName"]',document.body,null,9,null).singleNodeValue.value='{row['First Name']}';
        document.evaluate('//input[@ng-reflect-name="labelLastName"]',document.body,null,9,null).singleNodeValue.value='{row['Last Name']}';
        document.evaluate('//input[@ng-reflect-name="labelCompanyName"]',document.body,null,9,null).singleNodeValue.value='{row['Company Name']}';
        document.evaluate('//input[@ng-reflect-name="labelRole"]',document.body,null,9,null).singleNodeValue.value='{row['Role in Company']}';
        document.evaluate('//input[@ng-reflect-name="labelAddress"]',document.body,null,9,null).singleNodeValue.value='{row['Address']}';
        document.evaluate('//input[@ng-reflect-name="labelEmail"]',document.body,null,9,null).singleNodeValue.value='{row['Email']}';
        document.evaluate('//input[@ng-reflect-name="labelPhone"]',document.body,null,9,null).singleNodeValue.value='{row['Phone Number']}';  
        document.evaluate('//input[@value="Submit"]',document.body,null,9,null).singleNodeValue.click();      
    }}''')