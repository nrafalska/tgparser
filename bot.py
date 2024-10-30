import telebot
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from bs4 import BeautifulSoup
import os
import shutil
from datetime import datetime
import pdfkit  # Library for converting HTML to PDF

# Bot and site configuration
TELEGRAM_BOT_TOKEN = '7881751354:AAGJ0pmCTr5Lzrk57bbXljpcs9HqphZwMjM'
SITE_LOGIN_URL = 'https://app.creditrepaircloud.com/'
SITE_USERNAME = 'rafalskanastia@gmail.com'
SITE_PASSWORD = 'Ringo123'
login_url = "https://member.identityiq.com/"
dashboard_url = "https://member.identityiq.com/Dashboard.aspx"
credit_report_url = "https://member.identityiq.com/CreditReport.aspx"
bot = telebot.TeleBot(TELEGRAM_BOT_TOKEN)

# Path to wkhtmltopdf
config = pdfkit.configuration(wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe')

# Function to scrape client data from the site
def scrape_website_for_memo(driver, client_name):
    chrome_options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=chrome_options)
    print(f"Opening website: {SITE_LOGIN_URL}")
    driver.get(SITE_LOGIN_URL)

    username = SITE_USERNAME
    password = SITE_PASSWORD
    security_password = None

    try:
        print("Waiting for login field to appear...")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="username"]')))
        driver.find_element(By.XPATH, '//*[@id="username"]').send_keys(SITE_USERNAME)
        driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(SITE_PASSWORD)
        driver.find_element(By.XPATH, '//*[@id="signin"]').click()

        # Navigate to "Clients" page
        print("Waiting for Clients page to load...")
        WebDriverWait(driver, 10).until(EC.url_changes(SITE_LOGIN_URL))
        driver.get('https://app.creditrepaircloud.com/app/clients')

        print("Waiting for search field to appear...")
        search_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, 'tableSearch'))
        )
        search_field.send_keys(client_name)

        print(f"Waiting for search results for {client_name}...")
        time.sleep(5)

        # Searching for client
        print(f"Searching for client {client_name[:5]} in the list...")
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, "html.parser")

        found = False
        client_name_lower = client_name[:5].lower()
        for elem in soup.find_all(attrs={"title": True}):
            title_text = elem['title']
            title_text_lower = title_text.lower()
            if client_name_lower in title_text_lower:
                print(f"Found client with name {title_text}")
                href = elem.get('href', '')
                if href:
                    driver.get(href)
                    found = True
                    break

        if not found:
            print(f"Client with name {client_name[:5]} not found.")
            return username, password, security_password
        time.sleep(5)
        # Navigate to client's dashboard and extract security code
        print("Waiting for client's Dashboard page to load...")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'MuiTypography-root.MuiTypography-h1'))
        )
        dashboard_title = driver.find_element(By.CLASS_NAME, 'MuiTypography-root.MuiTypography-h1').text
        security_password = dashboard_title[-4:] if len(dashboard_title) >= 4 else "N/A"
        print(f"Security Code: {security_password}")

        # Navigate to Import/Audit tab
        print("Navigating to Import/Audit tab...")
        driver.find_element(By.XPATH, '//*[@id="root"]/div[2]/div[2]/div/div/div[3]/div[1]/div[2]/div/div/div/div[1]/div[2]/div[2]/button').click()
        time.sleep(5)

        # Extracting login and password
        try:
            print("Extracting login and password")
            username1 = driver.find_element(By.CLASS_NAME, 'username_lbl').text
            password1 = driver.find_element(By.CLASS_NAME, 'password_lbl').text
            print(f"Username: {username1}, Password: {password1}")
        except Exception as e:
            print(f"Error while extracting login and password: {e}")
            username1, password1 = None, None  # Assign None if not found
    finally:
        print("No browser closure here, browser stays open for further actions.")

    return username1, password1, security_password


# Function to log in to IdentityIQ and download the report
def login_to_identityiq(driver, username, password, security_password, chat_id):
    try:
        print(f"Logging in with username: {username}")
        time.sleep(5)

        # Use the same driver session from previous function
        driver.get("https://member.identityiq.com/")

        # Wait for Privacy Preference window and close it
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//button[contains(text(), "Confirm My Choices")]'))
            )
            driver.find_element(By.XPATH, '//button[contains(text(), "Confirm My Choices")]').click()
            print("Closed the Privacy Preference window")
            time.sleep(5)
        except Exception as e:
            print(f"Privacy preference window did not appear or could not be closed: {e}")

        # Enter login and password
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='txtUsername']")))
        driver.find_element(By.XPATH, "//*[@id='txtUsername']").send_keys(username)
        driver.find_element(By.XPATH, "//*[@id='txtPassword']").send_keys(password)
        driver.find_element(By.XPATH, "//*[@id='imgBtnLogin']").click()
        print("Login submitted")

        # Wait for security question page
        time.sleep(5)

        try:
            driver.find_element(By.XPATH, '//*[@id="FBfbforcechangesecurityanswer_txtSecurityAnswer"]').send_keys(security_password)
            driver.find_element(By.XPATH, '//*[@id="FBfbforcechangesecurityanswer_ibtSubmit"]').submit()
            print(f"Security code {security_password} submitted")
            WebDriverWait(driver, 20).until(
                EC.url_contains("Dashboard.aspx")
            )
            print("Successfully redirected to Dashboard.")
        except Exception as e:
            raise Exception("Error during security question handling")

        # Wait for Dashboard to load
        WebDriverWait(driver, 20).until(EC.url_to_be("https://member.identityiq.com/Dashboard.aspx"))
        print("Successfully logged into Dashboard.")
        time.sleep(5)

        # Navigate to report
        driver.get("https://member.identityiq.com/CreditReport.aspx")
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="reportTop"]/div[1]/div/div/a[2]'))
        )
        driver.find_element(By.XPATH, '//*[@id="reportTop"]/div[1]/div/div/a[2]').click()
        print("Downloading report...")

        time.sleep(5)  # Wait for the download to complete

        # Find the most recently downloaded file
        download_folder = "C:\\Users\\rumpelst\\Downloads"
        files = os.listdir(download_folder)
        files.sort(key=lambda x: os.path.getmtime(os.path.join(download_folder, x)), reverse=True)
        latest_file = os.path.join(download_folder, files[0])

        # Rename and move the file
        new_file_name = f"{username.replace('@', '_')}_{datetime.now().strftime('%Y-%m-%d')}.html"
        project_folder = os.getcwd()
        destination_folder = os.path.join(project_folder, "reports")
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
        shutil.move(latest_file, os.path.join(destination_folder, new_file_name))

        # Convert HTML to PDF
        html_file_path = os.path.join(destination_folder, new_file_name)
        pdf_file_path = os.path.join(destination_folder, new_file_name.replace('.html', '.pdf'))

        try:
            pdfkit.from_file(html_file_path, pdf_file_path, configuration=config)
            print(f"File converted: {new_file_name}")
        except Exception as e:
            raise Exception(f"Error converting file {new_file_name}: {str(e)}")

        # Send PDF to Telegram chat
        with open(pdf_file_path, 'rb') as pdf_file:
            bot.send_document(chat_id, pdf_file)
        print(f"Report sent to the user in Telegram: {new_file_name.replace('.html', '.pdf')}")
        time.sleep(5)

        driver.quit()

    except Exception as e:
        # Send error message to bot
        bot.send_message(chat_id, "Report for this client is unavailable")
        print(f"An error occurred: {e}")
        driver.quit()


# Handler for group messages where the bot is mentioned
@bot.message_handler(func=lambda message: message.chat.type in ['group', 'supergroup'])
def handle_group_message(message):
    # Check if the bot is mentioned
    if f"@{bot.get_me().username}" in message.text:
        try:
            # Extract the client name after the bot mention
            text_parts = message.text.split(f"@{bot.get_me().username}")
            if len(text_parts) > 1:
                client_name = text_parts[1].strip()
                chat_id = message.chat.id

                bot.reply_to(message, f"Searching for client: {client_name}. Please wait...")

                # Initialize WebDriver here and pass it through functions
                chrome_options = webdriver.ChromeOptions()
                driver = webdriver.Chrome(options=chrome_options)

                # Search for the client and get login details
                username, password, security_password = scrape_website_for_memo(driver, client_name)

                if username and password:
                    bot.reply_to(message, "Logging in and downloading the report...")
                    login_to_identityiq(driver, username, password, security_password, chat_id)
                else:
                    bot.reply_to(message, f"Client {client_name} not found.")
        except Exception as e:
            bot.reply_to(message, "Report for this client is unavailable")


# Start the bot
try:
    print("Bot started.")
    bot.polling(none_stop=True)
except Exception as e:
    print(f"Error: {e}")
