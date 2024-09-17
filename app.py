import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from flask import Flask, request, render_template
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

app = Flask(__name__)

# Gmail setup using environment variables
GMAIL_USER = os.getenv('GMAIL_USER')
GMAIL_PASSWORD = os.getenv('GMAIL_PASSWORD')
# TO_EMAILS = ['absk432@gmail.com']
TO_EMAILS = ['rrlcs1995@gmail.com']
# CC_EMAILS = ['sagar.singhal522@gmail.com', 'krishna617587@gmail.com']
CC_EMAILS = ['vrndavaneshwari@gmail.com']

# Google Sheets URL
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1-i3ZEiFQlnQJtDaWOp7zVf-iTFe_NPlAJqPT34bC0m4/edit?gid=107215885#gid=107215885'

# Function to take a screenshot of the Google Sheet
def take_screenshot():
    options = Options()
    options.headless = True
    service = ChromeService(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(SHEET_URL)

        # Wait for the page to load completely
        time.sleep(1)  # Wait for the page to load completely (adjust as necessary)

        # Take a screenshot of the entire visible sheet
        screenshot_path = 'static/sheet_screenshot.png'
        driver.save_screenshot(screenshot_path)
        return screenshot_path
    finally:
        driver.quit()

# Function to send email with the screenshot
def send_email_with_screenshot(screenshot_path):
    msg = MIMEMultipart()
    msg['Subject'] = 'Google Sheets Data Screenshot'
    msg['From'] = GMAIL_USER
    msg['To'] = ", ".join(TO_EMAILS)
    msg['Cc'] = ", ".join(CC_EMAILS)

    # Email body
    body = MIMEText('Please find the screenshot attached.', 'plain')
    msg.attach(body)

    # Attach the screenshot
    with open(screenshot_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(screenshot_path)}')
        msg.attach(part)

    to_emails = TO_EMAILS + CC_EMAILS

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(GMAIL_USER, GMAIL_PASSWORD)
            server.sendmail(GMAIL_USER, to_emails, msg.as_string())
        print('Email sent successfully!')
    except Exception as e:
        print(f'Failed to send email: {e}')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_screenshot', methods=['POST'])
def generate_screenshot():
    screenshot_path = take_screenshot()
    send_email_with_screenshot(screenshot_path)
    return render_template('index.html', screenshot_url=screenshot_path)

if __name__ == '__main__':
    app.run(debug=True)
