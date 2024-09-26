from selenium import webdriver
from selenium.webdriver.common.by import By
import csv
import time
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

s=Service(r"C:\Users\Admin\Desktop\21-09-2024_10-13\Backup 1\Single Source\Data Vault\Data Backup\Master_001\Master Folder\HD Backup\AIChorder.com\chromedriver-win64\chromedriver-win64\chromedriver.exe")

driver=webdriver.Chrome(service = s)

driver.get("https://www.linkedin.com/jobs/search/?currentJobId=4013790243&distance=25.0&f_JT=C&f_WT=2&geoId=103644278&keywords=Data%20Engineer&origin=JOB_SEARCH_PAGE_JOB_FILTER&sortBy=R")

# Wait for the page to load the initial job listings
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//ul[contains(@class, "jobs-search__results-list")]')))

scroll_pause_time = 2

while True:
    # Scroll down the page
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(scroll_pause_time)  # Wait for new listings to load

    # Check for the "See more jobs" button and click if available
    try:
        see_more_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'See more jobs')]"))
        )
        see_more_button.click()
        time.sleep(scroll_pause_time)  # Wait for more jobs to load
    except Exception as e:
        print("No more 'See more jobs' button:", e)
        break  # Exit the loop if no more button is found

# After loading all jobs, find all job elements
job_elements = driver.find_elements(By.XPATH, '//ul[contains(@class, "jobs-search__results-list")]/li')

# Open CSV file to write job details
with open('job_listings.csv', 'w', newline='', encoding='utf-8') as csvfile:
    csv_writer = csv.writer(csvfile)
    csv_writer.writerow(['Job Title', 'Company', 'Location'])

    # Iterate over each job element and extract details
    for listing in job_elements:
        try:
            job_title_element = listing.find_element(By.CLASS_NAME, 'base-search-card__title')
            job_title = job_title_element.text

            company_name_element = listing.find_element(By.CLASS_NAME, 'base-search-card__subtitle')
            company_name = company_name_element.text

            location_element = listing.find_element(By.CLASS_NAME, 'job-search-card__location')
            location = location_element.text

            print("Job Title:", job_title)
            print("Company:", company_name)
            print("Location:", location)
            print("-" * 50)

            csv_writer.writerow([job_title, company_name, location])

        except Exception as e:
            print("Error extracting job details:", e)

# Close the driver
driver.quit()