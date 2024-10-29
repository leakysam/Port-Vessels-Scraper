import os
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
from urllib.parse import urljoin

# Set up Chrome options
chrome_options = Options()
chrome_options.add_argument('--headless')  # Run in headless mode (no GUI)
chrome_options.add_argument('--disable-gpu')

# Set up the path to your ChromeDriver
driver_path = '/home/sam/Downloads/chromedriver-linux64 sept/chromedriver-linux64/chromedriver'  # Adjust this as necessary
service = Service(driver_path)

# Initialize the WebDriver
driver = webdriver.Chrome(service=service, options=chrome_options)

# Create a folder named 'Updated Ports' if it doesn't exist
folder_name = 'Updated Ports'
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# Get today's date, yesterday's date, and the day before yesterday's date in the required format
today_date = datetime.now().strftime("%d %b-%Y").upper()  # E.g., "01 OCT-2024"
yesterday_date = (datetime.now() - timedelta(days=1)).strftime("%d %b-%Y").upper()  # E.g., "30 SEP-2024"
day_before_yesterday_date = (datetime.now() - timedelta(days=2)).strftime("%d %b-%Y").upper()  # E.g., "29 SEP-2024"

# List of dates to check
dates_to_check = [today_date, yesterday_date, day_before_yesterday_date]

# Function to download Excel files from Mauport
def download_recent_excel_file(url):
    try:
        # Open the website
        driver.get(url)
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Extract the date from the specified span element
        date_element = soup.find('span', class_='date-display-single')
        if not date_element:
            print(f"No date element found on the page: {url}")
            return

        # Extract the content attribute, which contains the datetime
        upload_datetime_str = date_element['content']
        upload_datetime = datetime.fromisoformat(upload_datetime_str)
        upload_date = upload_datetime.date()

        # Check if the upload date is today or yesterday
        if upload_date not in (datetime.now().date(), (datetime.now() - timedelta(days=1)).date()):
            print(f'Skipped (not recent): Uploaded on {upload_date}')
            return

        # Find all links to Excel files
        excel_links = soup.find_all('a', href=lambda href: href and (href.endswith('.xlsx') or href.endswith('.xls')))
        if not excel_links:
            print(f"No Excel files found on the page: {url}")
            return

        for link in excel_links:
            # Get the full URL of the Excel file
            file_url = urljoin(url, link['href'])
            original_file_name = link.get_text(strip=True) + '.xlsx'
            original_file_name = original_file_name.replace('/', '_').replace('\\', '_')

            # Check if the file already exists in the folder
            if os.path.exists(os.path.join(folder_name, original_file_name)):
                print(f'Skipped (already exists): {original_file_name}')
            else:
                try:
                    file_response = requests.get(file_url, verify=False)
                    file_response.raise_for_status()

                    # Save the file in the 'Updated Ports' folder
                    with open(os.path.join(folder_name, original_file_name), 'wb') as f:
                        f.write(file_response.content)
                    print(f'Downloaded: {original_file_name} from {file_url}')
                except requests.exceptions.RequestException as e:
                    print(f'Failed to download: {original_file_name} - {e}')
    except Exception as e:
        print(f"An error occurred while downloading from {url}: {e}")

# Function to download Excel files from KPA
def download_kpa_excel_files(url):
    try:
        # Step 1: Open the webpage using Selenium
        driver.get(url)

        # Step 2: Wait for the download links to be present
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//a[contains(@href, ".xlsx")]')))

        # Step 3: Get the page source and parse with BeautifulSoup
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Step 4: Find the relevant links
        download_links = []
        for div in soup.find_all('div', class_='ms-vb'):
            link = div.find('a', href=True)
            if link and "MOMBASA 14 DAY LIST" in link.text:
                file_name = link['href'].split('/')[-1].replace('%20', ' ')  # Clean up spaces in the file name
                print(f"Checking file: {file_name}")

                # Check if the file matches one of the dates we're interested in
                if any(date in file_name for date in dates_to_check):
                    download_links.append(link['href'])

        # Step 5: Process found download links
        if download_links:
            for download_link in download_links:
                print(f'Found download link: {download_link}')

                # Extract the file name from the link
                file_name = download_link.split('/')[-1].replace('%20', ' ')

                # Check if the file already exists in the destination folder
                if os.path.exists(os.path.join(folder_name, file_name)):
                    print(f'Skipped (already exists): {file_name}')
                else:
                    try:
                        # Download the file if it doesn't already exist
                        file_response = requests.get(download_link, verify=False)
                        file_response.raise_for_status()

                        # Save the file in the 'Updated Ports' folder
                        with open(os.path.join(folder_name, file_name), 'wb') as f:
                            f.write(file_response.content)
                        print(f'Downloaded: {file_name}')
                    except requests.exceptions.RequestException as e:
                        print(f'Failed to download: {file_name} - {e}')
        else:
            print('No download link found for today, yesterday, or the day before yesterday.')

    except Exception as e:
        print(f"An error occurred while downloading from {url}: {e}")

# URLs to process
urls = [
    {'url': 'http://www.mauport.com/daily-port-situation/', 'function': download_recent_excel_file},
    {'url': 'https://www.kpa.co.ke/Pages/14DaysList.aspx?View={20888d2b-23c8-43a0-9cc6-a169bbb389c2}&SortField=Created&SortDir=Desc', 'function': download_kpa_excel_files}
]

# Loop over each URL and apply the corresponding function
for site in urls:
    site['function'](site['url'])

# Close the WebDriver after the process
driver.quit()
