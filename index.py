import glob
import os
import time
import os.path
import chromedriver_autoinstaller
from datetime import datetime
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# URL to the FanGraphs Auction Calculator to scrape data from
url = 'https://www.fangraphs.com/fantasy-tools/auction-calculator?teams=12&lg=MLB&dollars=260&mb=1&mp=20&msp=5&mrp=5&type=bat&players=&proj=thebatx&split=&points=c%7C0%2C1%2C2%2C3%2C4%7C13%2C14%2C2%2C3%2C4&rep=0&drp=0&pp=C%2CSS%2C2B%2C3B%2COF%2C1B&pos=1%2C1%2C1%2C1%2C4%2C1%2C0%2C0%2C1%2C2%2C4%2C2%2C4%2C6%2C0&sort=&view=0'

def main():
    print("Launching browser")
    launch_browser()
    print("Renaming file")
    fan_graph_file = rename_folder()
    print("Reading file")
    file_data = read_file(fan_graph_file)
    print("Setting up google sheets and transferring data")
    setup_google_sheets(file_data)

def launch_browser():

    try:
        # Set up code to run google chrome in headless mode and download the file
        opt = webdriver.ChromeOptions()
        download_path = str(Path.home() / "Downloads")
        opt.add_experimental_option("detach", False)
        opt.add_argument(f'download.default_directory={download_path}')
        opt.add_argument("--start-maximized")

        chromedriver_autoinstaller.install()
        driver = webdriver.Chrome(options=opt)
        driver.maximize_window() 
        driver.implicitly_wait(20)
        driver.get(url)

        fan_graph_download_btn = driver.find_element(By.CLASS_NAME, 'data-export')
        
        fan_graph_download_btn.click()

        time.sleep(2)

        driver.close()

    except Exception as e:
        print("Error has occured", e)

def setup_google_sheets(data):
    # Code checks if the token.json file exists and if it does not, it will create it
    # User will need to authenticate the app to access their google sheets data
    try:
        # If modifying these scopes, delete the file token.json.
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        """Shows basic usage of the Sheets API.
        Prints values from a sample spreadsheet.
        """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES
                )

                creds = flow.run_local_server(port=0)

            with open("token.json", "w") as token:
                token.write(creds.to_json())

        try:
            # Call the Sheets API and create a new spreadsheet
            service = build("sheets", "v4", credentials=creds)
            spreadsheet = {"properties": {"title": f"FanGraphs-Auction-Calculator-{datetime.today().strftime('%d-%m-%Y-%H-%M-%S')}", "locale": "en_US", "timeZone": "America/Los_Angeles"}}
            spreadsheet = (
                service.spreadsheets()
                .create(body=spreadsheet, fields="spreadsheetId")
                .execute()
            )

            sheet_id = spreadsheet.get("spreadsheetId")

            end_row = len(data)
            range_str = f'A1:A{end_row}'
            
            # Insert data into new google sheet
            service.spreadsheets().values().update(spreadsheetId=sheet_id, range="A1", valueInputOption="USER_ENTERED", body={"values": data}).execute()


            # Prepare to split the data in column A into multiple columns
            result = (
                    service.spreadsheets()
                    .values()
                    .get(spreadsheetId=sheet_id, range=range_str)
                    .execute()
                )
            
            rows = result.get("values", [])

            requests = []

            # Split the data in column A into multiple columns
            for row_num, row in enumerate(rows, start=1):
                forumla = f'=SPLIT(A{row_num}, ",")'
                cell = f'C{row_num}'
                requests.append({
                    'range': cell,
                    'values': [[forumla]],
                    'majorDimension': 'ROWS'
                })

            body = {
                'valueInputOption': 'USER_ENTERED',
                'data': requests
            }

            # Execute the request to split the data in column A into multiple columns
            result = service.spreadsheets().values().batchUpdate(spreadsheetId=sheet_id, body=body).execute()

            return spreadsheet.get("spreadsheetId")

        except HttpError as err:
            print(err)
        
    except Exception as e:
        print("Error on setup google sheets", e)

def rename_folder():
    # Rename the file to include the current date and time
    current_date = datetime.today().strftime('%d-%m-%Y')
    current_time = datetime.today().strftime('%H-%M-%S')
    home = os.path.expanduser('~')
    path = os.path.join(home, 'Downloads')

    path_a = path + "/*"
    list_of_files = glob.glob(path_a)

    latest_file = max(list_of_files, key=os.path.getctime)

    new_file = os.path.join(path, f"fangraphs-auction-calculator-{current_date}-{current_time}.xlsx")
    # Execute the rename
    os.rename(latest_file, new_file)

    return new_file

def read_file(file: str):
    items = []
    # Read the file and return the data inside the file
    with open(file, 'r') as f:
        lines = f.readlines()
        for item in lines:
            items.append([item.replace('\n', '')])
        f.close()
        return items
    

if __name__ == '__main__':
    main()