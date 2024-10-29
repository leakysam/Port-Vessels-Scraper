import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta

# Set up Chrome options
chrome_options = Options()
chrome_options.add_argument('--headless')  # Optional: Run in headless mode (no GUI)
chrome_options.add_argument('--disable-gpu')

# Set up the path to your ChromeDriver
driver_path = '/home/sam/Downloads/chromedriver-linux64 sept/chromedriver-linux64/chromedriver'  # Adjust this as necessary
service = Service(driver_path)

# Create a folder named 'Recent Updates' if it doesn't exist
folder_name = 'Recent Updates'
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# Get the current date and the previous date
current_date = datetime.now().date()
previous_date = current_date - timedelta(days=1)

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

        # Function to download the most recent Excel file from the current page (Mauport)
        def download_recent_excel_file():
            # Get the page source and parse with BeautifulSoup
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Extract the date from the specified span element
            date_element = soup.find('span', class_='date-display-single')
            if not date_element:
                print(f"No date element found on the page: {url}")
                return

            # Extract the content attribute, which contains the datetime
            upload_datetime_str = date_element['content']
            upload_datetime = datetime.fromisoformat(upload_datetime_str)  # Convert to datetime object
            upload_date = upload_datetime.date()  # Get just the date part

            # Check if the upload date is today or yesterday
            if upload_date not in (current_date, previous_date):
                print(f'Skipped (not recent): Uploaded on {upload_date}')
                return  # Skip if not the same day or the previous day

            # Find all links to Excel files
            excel_links = soup.find_all('a', href=lambda href: href and (href.endswith('.xlsx') or href.endswith('.xls')))
            if not excel_links:
                print(f"No Excel files found on the page: {url}")
                return

            # Assume only one file corresponds to the most recent date
            most_recent_file_url = None
            most_recent_file_name = None

            for link in excel_links:
                # Get the full URL of the Excel file
                file_url = urljoin(url, link['href'])  # Ensure the full URL is constructed correctly
                original_file_name = link.get_text(strip=True) + '.xlsx'  # Extract the file name

                # Clean the filename to remove illegal characters
                original_file_name = original_file_name.replace('/', '_').replace('\\', '_')

                # Update the most recent file details
                most_recent_file_url = file_url
                most_recent_file_name = original_file_name
                break  # Exit after finding the first link

            # Check if we found a file to download
            if most_recent_file_url and most_recent_file_name:
                # Check if the file already exists in the folder
                if os.path.exists(os.path.join(folder_name, most_recent_file_name)):
                    print(f'Skipped (already exists): {most_recent_file_name}')
                    return  # Skip downloading if the file already exists

                # Download the Excel file
                try:
                    file_response = requests.get(most_recent_file_url, verify=False)  # Bypassing SSL verification
                    file_response.raise_for_status()  # Raise an exception for HTTP errors

                    # Save the file in the 'Recent Updates' folder with its original name
                    with open(os.path.join(folder_name, most_recent_file_name), 'wb') as f:
                        f.write(file_response.content)
                    print(f'Downloaded: {most_recent_file_name} from {most_recent_file_url}')
                except requests.exceptions.RequestException as e:
                    print(f'Failed to download: {most_recent_file_name} - {e}')

        # Function to download Excel files from the current page (KPA)
        def download_excel_files():
            # Get the page source and parse with BeautifulSoup
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Find all links to Excel files using the provided HTML structure
            excel_links = soup.find_all('a', href=lambda href: href and (href.endswith('.xlsx') or href.endswith('.xls')))
            date_elements = soup.find_all('td', class_='ms-vb2')

            if not excel_links:
                print(f"No Excel files found on the page: {url}")
                return

            for link, date_element in zip(excel_links, date_elements):
                # Extract and parse the timestamp
                upload_time = date_element.get_text(strip=True)
                try:
                    # Parse the date and time from the element
                    upload_datetime = datetime.strptime(upload_time, '%m/%d/%Y %I:%M %p')  # Extract date and time
                    upload_date = upload_datetime.date()  # Get just the date part
                except ValueError as e:
                    print(f"Date parsing error for {upload_time}: {e}")
                    continue  # Skip if date parsing fails

                # Check if the upload date is today or yesterday
                if upload_date != current_date and upload_date != previous_date:
                    print(f'Skipped (not recent): {link.get_text(strip=True)} - Uploaded on {upload_date}')
                    continue  # Skip if not the same day or the previous day

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

                    # Save the file in the 'Recent Updates' folder with its original name
                    with open(os.path.join(folder_name, original_file_name), 'wb') as f:
                        f.write(file_response.content)
                    print(f'Downloaded: {original_file_name} from {file_url}')
                except requests.exceptions.RequestException as e:
                    print(f'Failed to download: {original_file_name} - {e}')

        # Conditional download based on the URL
        if url == 'http://www.mauport.com/daily-port-situation/':
            download_recent_excel_file()
        else:
            download_excel_files()

finally:
    # Close the WebDriver
    driver.quit()
