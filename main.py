from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By
import subprocess
import time

# Initialize Selenium WebDriver
# service = Service('path/to/chromedriver')  # Update with your WebDriver path
options = webdriver.ChromeOptions()
options.add_argument('--headless')  # Run in headless mode for automation

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

# driver = webdriver.Chrome(service=service, options=options)

try:
    # Run the Python script
    process = subprocess.Popen(['python', './reconciliation_script.py'], stdout=subprocess.PIPE,
                               stderr=subprocess.PIPE)
    stdout, stderr = process.communicate()

    # Check the output
    if "Reconciliation report generated" in stdout.decode():
        # Read the generated report
        with open('reconciliation_report.txt', 'r') as report:
            report_content = report.read()

        # You can extend this to email the report or upload it somewhere
        print(report_content)

    elif "Directories are in sync" in stdout.decode():
        print("Directories are in sync. No discrepancies found.")

    else:
        print("Error running the reconciliation script.")
        print(stderr.decode())

finally:
    driver.quit()
