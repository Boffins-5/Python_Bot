import threading
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC

# Global variables
driver_service = None
driver = None
google_sheet = None

def googleSheetsAccess():
    global google_sheet
    permissions = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    # in the path-to-your-key enter your credentials/ key in the format of json 
    pass_file =ServiceAccountCredentials.from_json_keyfile_name("path-to-your-key",permissions)
    client = gspread.authorize(pass_file)
    google_sheet = client.open("python").sheet1

def formAccess():
    chrome_driver = 'C:\\Program Files\\chromedriver-win64\\chromedriver.exe'
    return Service(chrome_driver)


def main():
    global driver
    googleSheetsAccess()
    formAccess()
    
    n = int(input("Enter the number of threads you want to start: "))
    threads = []
    data = google_sheet.get_all_records()
    
    driver = webdriver.Chrome(service=driver_service)
    i=0
    for row in data:
        if len(threads) < n:
            thread = threading.Thread(target=automate, args=(row.get('Name'), row.get('Email')))
            google_sheet.update(f'C{i+2}', [["done"]])  
            threads.append(thread)
            thread.start()
        else:
            time.sleep(2)
            threads = [] 
            thread = threading.Thread(target=automate, args=(row.get('Name'), row.get('Email')))
            google_sheet.update(f'C{i+2}', [["done"]])  
            threads.append(thread)
            thread.start()
        i+=1
    
    print("All forms submitted.")

def automate(name, email):
    driver_service = formAccess()
    driver = webdriver.Chrome(service=driver_service)    
    try:
        driver.get("https://tally.so/r/waDMG2")
        global button
        button = driver.find_element(By.TAG_NAME, "button")
        

        form_element_1 = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "e729bd5e-3362-4712-823c-9b426dcb0610"))
        )
        form_element_1.send_keys(name)
        
        form_element_2 = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "9271d54a-c70b-4375-ac4b-7ad4502d321d"))
        )
        form_element_2.send_keys(email)

        button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.TAG_NAME, "button"))
        )
        button.click()
        print(name,email)
        time.sleep(4) 
        driver.close()


    except Exception as e:
        print(f"Error processing {name}: {e}")
        
if __name__ == "__main__":
    
    main()
