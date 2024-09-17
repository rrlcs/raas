from apscheduler.schedulers.blocking import BlockingScheduler
import logging
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import datetime
from email.utils import formataddr
from PIL import Image
from email.mime.image import MIMEImage
import base64

# Setup logging
logging.basicConfig(filename='/home/cradlewise/raas/your_log.log', level=logging.DEBUG)

# Load environment variables from .env file
load_dotenv('/home/cradlewise/raas/.env')

# Gmail setup using environment variables
GMAIL_USER = os.getenv('GMAIL_USER')
GMAIL_PASSWORD = os.getenv('GMAIL_PASSWORD')
TO_EMAILS = ['absk432@gmail.com']
# TO_EMAILS = ['vrndavaneshwari@gmail.com']
CC_EMAILS = ['sagar.singhal522@gmail.com', 'krishna617587@gmail.com']
# CC_EMAILS = ['rrlcs1995@gmail.com']

# Google Sheets URL
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1-i3ZEiFQlnQJtDaWOp7zVf-iTFe_NPlAJqPT34bC0m4/edit?gid=107215885#gid=107215885'

FLAG_FILE = '/home/cradlewise/raas/stop_scheduler.flag'  # Flag file path

# def take_screenshot():
    # options = Options()
    # options.headless = True
    # service = ChromeService(executable_path=ChromeDriverManager().install())
    # driver = webdriver.Chrome(service=service, options=options)

    # try:
    #     driver.get(SHEET_URL)
    #     # Wait for the element to be present
    #     wait = WebDriverWait(driver, 20)
    #     element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="scrollable_right_0"]/div[2]/div[2]')))
    #     time.sleep(1)  # Adjust this as necessary
    #     screenshot_path = '/home/cradlewise/raas/sheet_screenshot.png'
    #     element.screenshot(screenshot_path)
    #     return screenshot_path
    # finally:
    #     driver.quit()

def take_screenshot():
    try:
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        driver.get(SHEET_URL)

        # Wait for the sheet to load
        time.sleep(1)

        # Locate the specific range, e.g., A1 to D7
        element = driver.find_element(By.XPATH, '//*[@id="scrollable_right_0"]/div[2]/div[2]')
        
        # Get the location and size of the element
        location = element.location
        location['x'] = 47
        location['y'] = 167
        size = element.size
        size['height'] = 192
        size['width'] = 405
        print(location, size, "location and size")

        # Take a full page screenshot
        driver.save_screenshot('full_screenshot.png')
        driver.quit()

        # Open the full screenshot and crop the specific range
        full_img = Image.open('full_screenshot.png')
        left = location['x']
        top = location['y']
        right = left + size['width']
        bottom = top + size['height']
        cropped_img = full_img.crop((left, top, right, bottom))

        screenshot_path = 'screenshot.png'
        cropped_img.save(screenshot_path)

        return screenshot_path
    except Exception as e:
        logging.error(f"Error taking screenshot: {e}")
        if driver:
            driver.quit()
        raise

def send_email_with_screenshot(screenshot_path):
    try:
        # Get the current date
        current_date = datetime.date.today()

        # Calculate the start and end date of last week
        start_date = current_date - datetime.timedelta(days=current_date.weekday() + 7)
        end_date = current_date - datetime.timedelta(days=current_date.weekday() + 1)

        # Format the dates as required
        start_date_str = start_date.strftime('%Y-%m-%d')
        end_date_str = end_date.strftime('%Y-%m-%d')

        # Create the email message
        msg = MIMEMultipart('related')
        msg['Subject'] = f'Sadhana Card for {start_date_str} to {end_date_str}'
        msg['From'] = formataddr(("Ravi Raja", GMAIL_USER))
        msg['To'] = ", ".join(TO_EMAILS)
        msg['Cc'] = ", ".join(CC_EMAILS)

        # Create the body with the image embedded
        body = MIMEMultipart('alternative')
        html_body = f"""
        <html>
        <body>
            <p>Hare Krishna Prabhuji,<br>
            Please accept my humble obeisances.<br>
            All glories to guru and gauranga.<br>
            All glories to Srila Prabhupada.<br>
            <br>
            Please find attached the Sadhana Card for the previous week.<br>
            <br>
            <img src="cid:screenshot"><br>
            <br>
            Your servant,<br>
            Ravi Raja
            </p>
        </body>
        </html>
        """
        body.attach(MIMEText(html_body, 'html'))
        msg.attach(body)

        # Attach the image
        with open(screenshot_path, 'rb') as f:
            img_data = f.read()
        image = MIMEImage(img_data)
        image.add_header('Content-ID', '<screenshot>')
        image.add_header('Content-Disposition', 'inline', filename=os.path.basename(screenshot_path))
        msg.attach(image)

        to_emails = TO_EMAILS + CC_EMAILS

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(GMAIL_USER, GMAIL_PASSWORD)
            server.sendmail(GMAIL_USER, to_emails, msg.as_string())
        logging.info('Email sent successfully!')
    except Exception as e:
        logging.error(f'Failed to send email: {e}')

logging.info('Job started')
screenshot_path = take_screenshot()
send_email_with_screenshot(screenshot_path)
logging.info('Job finished')

# def job():
#     logging.info('Job started')
#     screenshot_path = take_screenshot()
#     send_email_with_screenshot(screenshot_path)
#     logging.info('Job finished')

# scheduler = BlockingScheduler()
# # scheduler.add_job(job, 'interval', minutes=1)
# scheduler.add_job(job, 'cron', day_of_week='tue', hour=21, minute=5)
# scheduler.start()
# # sp = take_screenshot()