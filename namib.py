import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from datetime import datetime

# Suppress InsecureRequestWarning from urllib3
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Create a folder named 'Updated Ports' if it doesn't exist
folder_name = 'Updated Ports'
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# The base URL of the website
base_url = 'https://www.namport.com.na/'

# The URL of the page containing the Excel file links
url = 'https://www.namport.com.na/ports/port-of-walvis-bay/548/'

# Fetch the page content
response = requests.get(url, verify=False)  # Bypassing SSL verification
response.raise_for_status()  # Raise an exception for HTTP errors

# Parse the HTML content with BeautifulSoup
soup = BeautifulSoup(response.content, 'html.parser')

# Find all relevant links to Excel files
excel_links = soup.find_all('a', href=lambda href: href and 'files/files/' in href)

# Store filenames for the files that are already downloaded
existing_files = {f for f in os.listdir(folder_name) if f.endswith('.xlsx')}

# Get the most recent file date from the existing files
recent_date = None
for existing_file in existing_files:
    try:
        date_str = existing_file.replace('_', ' ').split(' ')[0:3]  # e.g., "02 September 2024"
        file_date = datetime.strptime(' '.join(date_str), '%d %B %Y').date()
        if recent_date is None or file_date > recent_date:
            recent_date = file_date
    except ValueError:
        continue  # Skip if the file name format is unexpected

# Iterate through each link to download if not already present
for link in excel_links:
    # Extract the date from the span element or surrounding text
    date_span = link.find('span', style='font-size:11px')
    if date_span:
        upload_date_str = date_span.get_text(strip=True)
        upload_date = datetime.strptime(upload_date_str, '%d %B %Y').date()

        # Format the filename
        original_file_name = f"{upload_date_str.replace(' ', '_')}.xlsx"  # Generate the file name
    else:
        # If no date found, skip this link
        continue

    # Join the base URL with the relative URL from the href
    file_url = urljoin(base_url, link['href'])

    # Check if the file already exists in the folder
    if original_file_name in existing_files:
        print(f'Skipped (already exists): {original_file_name}')
        continue  # Skip downloading if the file already exists

    # Only download if the date is more recent than the most recent downloaded file
    if recent_date and upload_date <= recent_date:
        print(f'Skipped (not recent): {original_file_name}')
        continue  # Skip if the file is not more recent than the last downloaded file

    # Download the Excel file
    try:
        file_response = requests.get(file_url, verify=False)  # Bypassing SSL verification
        file_response.raise_for_status()  # Raise an exception for HTTP errors

        # Save the file in the 'Updated Ports' folder with its original name
        with open(os.path.join(folder_name, original_file_name), 'wb') as f:
            f.write(file_response.content)
        print(f'Downloaded: {original_file_name} from {file_url}')
    except requests.exceptions.RequestException as e:
        print(f'Failed to download: {original_file_name} - {e}')
