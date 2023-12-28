import csv
import os

from openpyxl import Workbook
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import time


excel_file_path = os.path.join(os.getcwd(), 'Test Ayoub New.xlsx')

# Read the Excel file
df = pd.read_excel(excel_file_path, dtype=str)  # Force all data to be read as string
# Extract the lines into a dictionary grouped by Boutique
data_grouped_by_boutique = df.groupby('Boutique')

# Placeholder for web scraping results
web_scraped_data = []

# Define Chrome options
chrome_options = Options()
chrome_options.add_argument('--headless')  # Run Chrome in headless mode (no GUI)
chrome_options.add_argument("--user-data-dir=C:\\Users\\Aouub\\AppData\\Local\\Google\\Chrome\\User Data")
chrome_options.add_argument("--profile-directory=Default")

# Setting up the WebDriver using ChromeDriverManager
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# URL to scrape
url = 'https://www.aliexpress.com/p/order/index.html?spm=a2g0o.home.headerAcount.2.2f5c2145tr74tz'

try:
    print("Opening URL...")
    driver.get(url)
    wait = WebDriverWait(driver, 10)

    print("Clicking on 'Shipped' tab...")
    shipped_div = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'comet-tabs-nav-item') and contains(text(), 'Shipped')]")))
    shipped_div.click()

    def parse_date(text):
        for part in text.split('\n'):
            if 'Order date:' in part:
                date_str = part.split('Order date: ')[1]
                return datetime.strptime(date_str, '%b %d, %Y')
        return None

    fifteen_days_ago = datetime.now() - timedelta(days=13)

    while True:
        print("Fetching order items...")
        order_items = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'order-item')))
        last_order_date = parse_date(order_items[-1].find_element(By.CLASS_NAME, 'order-item-header-right-info').text)
        print(f"Last order date: {last_order_date}")

        if last_order_date and last_order_date < fifteen_days_ago:
            print("Last order date is more than 15 days ago, breaking loop...")
            break

        print("Clicking 'More' button...")
        more_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'order-more')))
        more_button.click()
        time.sleep(5)  # Adjust sleep time as needed

    print("Collecting links...")
    button_elements = driver.find_elements(By.CLASS_NAME, 'comet-btn.comet-btn-block.order-item-btn')
    links = [element.get_attribute('href') for element in button_elements]
    print(f"Collected {len(links)} links.")

    tracking_numbers = []

    for link in links:
        if link:
            print(f"Navigating to {link}...")
            driver.get(link)
            try:
                print("Fetching tracking number and AliExpress Order ID...")
                tracking_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'tracking-no')))
                tracking_number = tracking_element.text
                print(f"Found tracking number: {tracking_number}")

                order_id_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'order-id')))
                order_id_span = order_id_element.find_element(By.CLASS_NAME, 'value')
                aliexpress_order_id = order_id_span.text
                print(f"Found AliExpress Order ID: {aliexpress_order_id}")

                web_scraped_data.append(
                    {'tracking_number': tracking_number, 'aliexpress_order_id': aliexpress_order_id})
            except Exception as e:
                print(f"Error fetching data for link {link}: {e}")

finally:
    print("Closing WebDriver...")
    driver.quit()

date_string = datetime.now().strftime('%Y-%m-%d')

general_data = []

# Process the matching data for each boutique
for boutique, data in data_grouped_by_boutique:
    matching_data = []
    for _, row in data.iterrows():
        wordpress_order_id, aliexpress_order_id = row['Order ID Wordpress'], row['Order ID AliExpress']
        for scraped in web_scraped_data:
            if aliexpress_order_id == scraped['aliexpress_order_id']:
                matching_data.append((wordpress_order_id, scraped['tracking_number'], 'cainiao'))

    # Check the number of commands for each boutique
    if len(matching_data) >= 5:
        # Create a CSV file for boutiques with 5 or more commands
        csv_file_name = f'{date_string}_{boutique}_order.csv'
        with open(csv_file_name, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Order ID", "Tracking Number", "Carrier Slug"])  # Write headers
            writer.writerows(matching_data)
        print(f"Data saved to {csv_file_name}")
    else:
        # Append to general_data for boutiques with less than 5 commands
        for entry in matching_data:
            general_data.append(entry + (boutique,))

# Create a single Excel file for all boutiques with less than 5 commands
if general_data:
    wb = Workbook()
    ws = wb.active
    ws.append(["Order ID", "Tracking Number", "Carrier Slug", "Boutique"])  # Write headers
    for data in general_data:
        ws.append(data)
    general_excel_file_name = f'{date_string}_general_order.xlsx'
    wb.save(general_excel_file_name)
    print(f"Data saved to {general_excel_file_name}")



