import selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

def verify_xpath(url, xpath):
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = selenium.webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    try:
        driver.get(url)
        
        # Give the page some time to load
        time.sleep(10)

        # Switch to the frame containing the sheet if necessary
        # Uncomment and adjust if the element is inside an iframe
        # iframe = driver.find_element(By.XPATH, 'iframe_xpath')
        # driver.switch_to.frame(iframe)

        # Wait for the element to be present
        wait = WebDriverWait(driver, 20)
        element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))

        if element:
            print("XPath is correct. Element found!")
        else:
            print("XPath is incorrect. Element not found.")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        driver.quit()

# Example usage
verify_xpath('https://docs.google.com/spreadsheets/d/1-i3ZEiFQlnQJtDaWOp7zVf-iTFe_NPlAJqPT34bC0m4/edit', '//*[@id="107215885-grid-table-container"]')
