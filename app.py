import time
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Step 1: Set up the Brave browser with Selenium
options = Options()

# No headless argument, so the browser will be visible
options.add_argument('--disable-gpu')  # Optional but useful
options.add_argument('--no-sandbox')  # Optional for some environments

# Path to chromedriver.exe (ensure chromedriver is in the same directory as the script)
chromedriver_path = 'chromedriver-win64/chromedriver.exe'  # Ensure chromedriver is in the same directory as the script

# Initialize WebDriver
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service, options=options)

# Base URL for the login page
login_url = "https://mrecacademics.com/"

# Function to login and fetch data
def login_and_fetch_data(username, password):
    driver.get(login_url)
    print(f"Navigating to login page: {login_url}")

    try:
        # Wait for username field to be present
        username_field = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="UserName"]'))
        )
        username_field.send_keys(username)
        print("Entered username.")

        # Wait for password field to be present
        password_field = driver.find_element(By.XPATH, '//*[@id="Password"]')
        password_field.send_keys(password)
        print("Entered password.")

        # Submit the login form
        password_field.send_keys(Keys.RETURN)
        print(f"Logging in with username: {username}")

        # Wait for page redirection and ensure the new page has loaded (use URL or an element that confirms redirection)
        WebDriverWait(driver, 15).until(
            EC.url_changes(login_url)  # Wait for the URL to change after login
        )

        # Optionally, check the current URL to confirm redirection (debugging step)
        print("Page redirected to:", driver.current_url)

        # Wait for dashboard to load (ensure we are on the correct page)
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="StudentName"]'))  # Wait for StudentName to appear
        )
        print("Login successful! Now scraping the data...")

    except Exception as e:
        print(f"Error during login: {e}")
        return {
            'StudentName': 'Error',
            'ParentMobile': 'Error',
            'Username': username,
            'Error': str(e)
        }

    # Step 4: Scrape data (specifically StudentName and ParentMobile)
    try:
        # Scraping StudentName (value of the input field)
        student_name_element = driver.find_element(By.XPATH, '//*[@id="StudentName"]')
        student_name = student_name_element.get_attribute('value')  # Get value attribute
        print(f"Student Name: {student_name}")

        # Scraping ParentMobile (value of the input field)
        parent_mobile_element = driver.find_element(By.XPATH, '//*[@id="ParentMobile"]')
        parent_mobile = parent_mobile_element.get_attribute('value')  # Get value attribute
        print(f"Parent Mobile: {parent_mobile}")

        # Store the extracted data in a dictionary
        data = {
            'StudentName': student_name,
            'ParentMobile': parent_mobile,
            'Username': username,
            'Error': None
        }

        return data

    except Exception as e:
        print(f"Error during data scraping: {e}")
        return {
            'StudentName': 'Error',
            'ParentMobile': 'Error',
            'Username': username,
            'Error': str(e)
        }

# Step 5: Main Execution - Fetch data for a list of users
data_list = []

for i in range(6701, 6710):  # Iterating from 6701 to 6799
    username = f"22J41A{i}"  # Username format: 22J41A6701 to 22J41A6799
    password = username  # Password is the same as the username
    print(f"Fetching data for {username}...")

    data = login_and_fetch_data(username, password)
    
    # Check if the data was successfully returned and add to the list
    if data['StudentName'] != 'details wrong':
        print(f"Data collected for {username}: {data}")
    else:
        print(f"Failed to collect data for {username}: {data['Error']}")

    data_list.append(data)  # Append each user's data to the list
    time.sleep(2)  # Optional: wait a bit between requests to prevent being blocked

# Step 6: Save the data to an Excel file using pandas
file_name = 'dashboard_data.xlsx'

try:
    # Check if data_list contains any data
    if len(data_list) > 0:
        # Create a DataFrame from the list of data dictionaries
        df = pd.DataFrame(data_list)

        # Check if the file exists
        if not os.path.exists(file_name):
            print(f"File '{file_name}' does not exist, creating it now.")
            df.to_excel(file_name, index=False)
            print(f"Data has been saved to '{file_name}'")
        else:
            # If the file exists, append the new data to it
            print(f"File '{file_name}' exists, appending data.")
            existing_df = pd.read_excel(file_name)  # Read existing data
            updated_df = existing_df.append(df, ignore_index=True)  # Append new data
            updated_df.to_excel(file_name, index=False)  # Save the updated data
            print(f"Data has been appended to '{file_name}'")
    else:
        print("No data collected. No file will be created.")
except Exception as e:
    print(f"Error saving data to Excel: {e}")

# Step 7: Close the browser
try:
    driver.quit()
    print("Browser closed successfully.")
except Exception as e:
    print(f"Error closing the browser: {e}")
