from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def navigate_to_page(browser, page_number):
    for _ in range(page_number - 1):  # Assuming you start on page 1
        try:
            next_button = browser.find_element(By.XPATH, "//a[@class='next']")
            browser.execute_script("arguments[0].click();", next_button)
            time.sleep(3)  # Static wait for the page to load
        except NoSuchElementException:
            print("Next button not found, cannot navigate further.")
            break

def scrape_and_save_to_sheet():
    # Set up the Selenium WebDriver with headless options
    options = Options()
    options.add_argument("--headless")  # Run in headless mode
    service = Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(service=service, options=options)

    # Google Sheets interaction setup
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'E:\PycharmProjects\Stocks\stockautomation-408107-5d32eaf884c1.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open('stocks_automation').sheet1

    try:
        url = 'https://www.nirmalbang.com/equity/highest-lowest-delivery.aspx'
        browser.get(url)
        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "modality")))

        # Navigate to page 20 before starting the scraping loop
        # start_page = 71
        # navigate_to_page(browser, start_page)
        #
        # page_count = start_page - 1
        page_count = 0
        while True:
            page_count += 1
            print(f"Scraping page {page_count}")

            rows = browser.find_elements(By.XPATH, '//*[@id="modality"]/tbody/tr')
            page_data = []
            for row in rows:
                columns = row.find_elements(By.TAG_NAME, 'td')
                row_data = [col.text for col in columns]
                page_data.append(row_data)

            # Save the scraped data of the current page to Google Sheets
            for row in page_data:
                sheet.append_row(row)
            print(f"Data from page {page_count} saved to Google Sheets")

            try:
                # Find the 'Next' button for each new page, as the old reference becomes stale
                next_button_li = browser.find_element(By.XPATH, "//a[@class='next']/ancestor::li")
                if 'disabled' in next_button_li.get_attribute('class'):
                    print("Reached the last page, ending scraping")
                    break

                next_button = browser.find_element(By.XPATH, "//a[@class='next']")
                browser.execute_script("arguments[0].click();", next_button)

                # Wait for some time after clicking the 'Next' button
                time.sleep(10)  # Static wait for 10 seconds

                print("Clicked 'Next' button and waiting for the new page")
            except (NoSuchElementException, TimeoutException):
                print("No 'Next' button found or timed out, ending scraping")
                break

        print("Data scraping and saving completed")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        browser.quit()

if __name__ == "__main__":
    result = scrape_and_save_to_sheet()
    print(result)
