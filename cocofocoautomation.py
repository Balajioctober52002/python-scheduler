import os
import time
import shutil
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# --- Web Scraping Functions ---

def setup_chrome_driver(download_directory):
    """Sets up Chrome WebDriver with a custom download directory."""
    from selenium.webdriver.chrome.options import Options

    chrome_options = Options()
    prefs = {
        "download.default_directory": download_directory,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    return driver

def login_to_website(driver, username, password):
    """Logs into the website."""
    try:
        driver.get('https://dms.mytvs.in/tvsfit/users/login')
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//input[@name="data[User][username]"]'))
        ).send_keys(username)

        driver.find_element(By.ID, 'UserPassword').send_keys(password)
        driver.find_element(By.CLASS_NAME, "btn-primary").click()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//body[contains(@class, "sidebar-mini")]'))
        )
        print("Login successful!")
    except Exception as e:
        print(f"Error during login: {e}")

def download_reports(driver, download_dir):
    """Navigates the menu, selects date ranges, and downloads the reports."""
    try:
        # Navigate to the menu
        treeview_item = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "(//li[@class='treeview'])[2]"))
        )
        ActionChains(driver).move_to_element(treeview_item).click().perform()

        link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "(//ul[@class='treeview-menu'])[2]//a"))
        )
        link.click()

        # Calculate date range
        today = datetime.today()
        first_of_month = today.replace(day=1).strftime('%d-%m-%Y')  # First date of the month
        yesterday = (today - timedelta(days=1)).strftime('%d-%m-%Y')  # Yesterday's date

        # Select date range
        date_picker = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'PartsIndentDaterange'))
        )
        date_picker.click()

        # Enter start date (first day of the month)
        start_date_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "(//input[@name='daterangepicker_start'])[1]"))
        )
        start_date_input.clear()
        start_date_input.send_keys(first_of_month)

        # Enter end date (yesterday's date)
        end_date_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "(//input[@name='daterangepicker_end'])[1]"))
        )
        end_date_input.clear()
        end_date_input.send_keys(yesterday)

        # Apply the date range
        driver.find_element(By.CLASS_NAME, 'applyBtn').click()

        # Trigger file download (first file)
        driver.find_element(By.ID, 'myButton').click()
        time.sleep(20)  # Wait for the first download to complete

        # Download second file (switch dropdown)
        driver.find_element(By.CLASS_NAME, "select2-selection--single").click()
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "select2-results__option"))
        )[1].click()
        driver.find_element(By.ID, 'myButton').click()
        time.sleep(20)  # Wait for the second download to complete

        print("Reports downloaded successfully!")
    except Exception as e:
        print(f"Error during report download: {e}")


def rename_downloaded_files(download_dir):
    """Renames the downloaded files to identifiable names."""
    today_date = datetime.today().strftime('%Y-%m-%d')
    foco_file = os.path.join(download_dir, "Sales_Gross_Margin.xlsx")
    coco_file = os.path.join(download_dir, "Sales_Gross_Margin (1).xlsx")

    if os.path.exists(foco_file):
        foco_renamed = os.path.join(download_dir, f"FOCO_{today_date}.xlsx")
        os.rename(foco_file, foco_renamed)
        print(f"Renamed {foco_file} to {foco_renamed}")
    if os.path.exists(coco_file):
        coco_renamed = os.path.join(download_dir, f"COCO_{today_date}.xlsx")
        os.rename(coco_file, coco_renamed)
        print(f"Renamed {coco_file} to {coco_renamed}")

    return foco_renamed, coco_renamed

# --- Processing Functions ---

def process_file(file_path, division_value, output_file_path):
    """Adds a 'Division' column to the Excel file."""
    df = pd.read_excel(file_path)
    df.insert(df.columns.get_loc('Branch') + 1, 'Division', division_value)
    df.to_excel(output_file_path, index=False)
    print(f"Processed file saved at: {output_file_path}")

def merge_files(file1_path, file2_path, output_file_path):
    """Merges two Excel files."""
    df1 = pd.read_excel(file1_path)
    df2 = pd.read_excel(file2_path)
    merged_df = pd.concat([df1, df2], ignore_index=True)
    merged_df.to_excel(output_file_path, index=False)
    print(f"Merged file saved at: {output_file_path}")

# --- Email Functions ---

def send_email_with_attachment(sender_email, app_password, recipient_email, cc_emails, attachment_path):
    """Sends an email with the merged file attached."""
    subject = "Merged Excel File"
    body = "Hello,\n\nPlease find the merged Excel file attached.\n\nBest regards,\nYour Name"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Cc'] = ', '.join(cc_emails)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_path, "rb") as file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment_path)}"')
    msg.attach(part)

    all_recipients = [recipient_email] + cc_emails
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(sender_email, app_password)
        server.sendmail(sender_email, all_recipients, msg.as_string())
    print("Email sent successfully!")

# --- Main Script ---

if __name__ == "__main__":
    download_dir = r"C:\Ki-Intern\Vishal\downloads"
    os.makedirs(download_dir, exist_ok=True)

    # Web scraping
    driver = setup_chrome_driver(download_dir)
    try:
        login_to_website(driver, 'K01068', 'Tvs@12349')
        download_reports(driver, download_dir)
    finally:
        driver.quit()

    # File renaming and processing
    foco_file, coco_file = rename_downloaded_files(download_dir)
    processed_foco = r"C:\Ki-Intern\FOCO_Processed2.xlsx"
    processed_coco = r"C:\Ki-Intern\COCO_Processed2.xlsx"
    process_file(foco_file, 'FOCO', processed_foco)
    process_file(coco_file, 'COCO', processed_coco)

    # File merging
    merged_file = r"C:\Ki-Intern\Merged_File(Coco & Foco-15th dec.xlsx"
    merge_files(processed_coco, processed_foco, merged_file)

    # Email sending
    send_email_with_attachment(
        sender_email="balajioctober52002@gmail.com",
        app_password="vxtdwuhtnrpchlyo",
        recipient_email="mackson.prince@tvs.in",
        cc_emails=[ "vishalsaranath@gmail.com", "gkmuralidharan1970@gmail.com","dheepesh.karthik@tvs.in","amogh.mavanthoor@tvs.in","karuppusamy.jayaraj@tvs.in","nithya.murugaiyan@tvs.in"],
        attachment_path=merged_file
    )
