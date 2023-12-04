import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

# Function to print the progress bar
def print_progress_bar(iteration, total, prefix='', suffix='', length=50, fill='█'):
    percent = ("{0:.1f}").format(100 * (iteration / float(total)))
    filled_length = int(length * iteration // total)
    bar = fill * filled_length + '-' * (length - filled_length)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end='\r')
    if iteration == total: 
        print()

# Chrome options for headless mode
chrome_options = Options()
chrome_options.add_argument("--headless")

# Load data from Excel file
excel_path = "data.xlsx"  # Replace with your file path
df = pd.read_excel(excel_path)

# Setting up Selenium WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 10)  # Adjust the wait time as necessary

# Function to visit each row link and scrape data
def visit_and_scrape(row):
    try:
        # Visit the link
        driver.get(row['Profile'])  # Assuming 'Profile' is the column name with URLs

        # Handling popup if it appears
        try:
            button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/app-login/div/div/div/div/div[3]/p/app-basic-button/button')))
            button.click()
        except TimeoutException:
            pass  # No popup appeared or timeout occurred

        # Scrape required data
        product_info = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#main-container .padding-half--top'))).text
        company_info = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.ng-trigger-fadeIn:nth-child(2) .padding--bottom'))).text

        # Return the scraped data
        return product_info, company_info

    except Exception as e:
        print(f"Error occurred: {e}")
        return None, None

# Iterating over each row in the DataFrame
batch_size = 100
total_rows = len(df)
for index, row in df.iterrows():
    product, company_info = visit_and_scrape(row)
    df.at[index, 'Product'] = product
    df.at[index, 'Company Info'] = company_info

    # Optional: pause between requests
    time.sleep(2)

    # Update progress bar
    print_progress_bar(index + 1, total_rows, prefix='Progress:', suffix='Complete', length=50)

    # Save in batches of 100
    if (index + 1) % batch_size == 0 or (index + 1) == total_rows:
        batch_number = (index + 1) // batch_size
        batch_file_name = f"updated_data_batch_{batch_number}.xlsx"
        df.iloc[max(index + 1 - batch_size, 0):index + 1].to_excel(batch_file_name, index=False)
        print(f"Batch {batch_number} saved to {batch_file_name}")

# Closing the WebDriver
driver.quit()

print("Scraping completed. All data batches saved.")
