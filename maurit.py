import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Set up Chrome options
chrome_options = Options()
chrome_options.add_argument('--headless')  # Optional: Run in headless mode (no GUI)
chrome_options.add_argument('--disable-gpu')

# Set up the path to your ChromeDriver
driver_path = '/home/sam/Downloads/chromedriver-linux64 (1)/chromedriver-linux64/chromedriver'  # Adjust this as necessary
service = Service(driver_path)

# Create a folder named 'KPA' if it doesn't exist
folder_name = 'mauritius'
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# Initialize the WebDriver
driver = webdriver.Chrome(service=service, options=chrome_options)

# List of URLs to process
urls = [
    'http://www.mauport.com/daily-port-situation/'
]

try:
    for url in urls:
        # Open the website
        driver.get(url)

        # Function to download Excel files from the current page
        def download_excel_files():
            # Get the page source and parse with BeautifulSoup
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Find all links to Excel files using the provided HTML structure
            excel_links = soup.find_all('a', href=lambda href: href and (href.endswith('.xlsx') or href.endswith('.xls')))

            if not excel_links:
                print(f"No Excel files found on the page: {url}")
                return

            for link in excel_links:
                # Get the full URL of the Excel file
                file_url = urljoin(url, link['href'])  # Ensure the full URL is constructed correctly
                original_file_name = link.get_text(strip=True) + '.xlsx'  # Extract the file name

                # Clean the filename to remove illegal characters
                original_file_name = original_file_name.replace('/', '_').replace('\\', '_')

                # Check if the file already exists in the folder
                if os.path.exists(os.path.join(folder_name, original_file_name)):
                    print(f'Skipped (already exists): {original_file_name}')
                    continue  # Skip downloading if the file already exists

                # Download the Excel file
                try:
                    file_response = requests.get(file_url, verify=False)  # Bypassing SSL verification
                    file_response.raise_for_status()  # Raise an exception for HTTP errors

                    # Save the file in the 'KPA' folder with its original name
                    with open(os.path.join(folder_name, original_file_name), 'wb') as f:
                        f.write(file_response.content)
                    print(f'Downloaded: {original_file_name} from {file_url}')
                except requests.exceptions.RequestException as e:
                    print(f'Failed to download: {original_file_name} - {e}')

        # Download files from the current URL
        download_excel_files()

finally:
    # Close the WebDriver
    driver.quit()
