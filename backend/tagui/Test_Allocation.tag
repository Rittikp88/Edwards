py begin

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import logging
import datetime

# Set up logging configuration
logging.basicConfig(filename='assignment_log.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

#app password
"pjwu glcx odev lfgy"

# Email configuration
sender_email = 'ashu.au55555.com'  # Replace with your email address
sender_password = 'pjwu glcx odev lfgy'        # Replace with your email password
smtp_server = 'smtp.gmail.com'         # Replace with your SMTP server

# Replace 'your_excel_file.xlsx' with the actual Excel file path
excel_file_path = 'D:\Edward\Mapped data.xlsx'

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(excel_file_path)

# Create a new instance of the Chrome webdriver
driver = webdriver.Chrome()

# Replace 'your_url' with the actual URL you want to open
url = 'https://himanshunegi378.github.io/saumya/'

# Open the URL in the browser
driver.get(url)

# Iterate over each row and print the data
for index, row in df.iterrows():
    
    materials = row['materials']
    studies = row['studies']
    protocols = row['protocols']
    test = row['test']
    resources = row['resources']
    email_column = row['email id']
    
    #log file Resources not available
    
    # Construct the XPath expression dynamically
    material_expression = f"//h2[text()='{materials}']" 
    studie_expression = f"//h2[text()='{studies}']" 
    protocol_expression = f"//h2[text()='{protocols}']" 
    test_expression = f"//h2[text()='{test}']" 
    resource_expression = f"//h2[text()='{resources}'][following-sibling::button[@class='chakra-button css-10uk6k7' and text()='Available']]" 
    allocate_button = f"/html/body/div[1]/div/div[2]/div/div[5]/div/button" 
    no_resource_message = "No Resource Available"

     # Check if "No Resource Available" message is present
    if no_resource_message in resources:
        logging.error(f"no resource available in row: {index+1}")
        logging.error(f"Error in row {index+1} : {materials},{studies},{protocols},{test},{resources}")
        continue  # Move to the next row
    # Create a message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = email_column
    message['Subject'] = 'Task Assignment Details'

    # Body of the email
    body = f"Hello {resources},\n\nYou have been assigned a task and here are the details:\n"
    body += f"Material: {materials}\n"
    body += f"Study: {studies}\n"
    body += f"Protocol: {protocols}\n"
    body += f"Test: {test}\n\nThank you"

    # Attach the body to the message
    message.attach(MIMEText(body, 'plain'))
    
    try:
        # Wait up to 10 seconds for the element to be present

        # Material List
        element1 = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, material_expression))
        )
        time.sleep(1)
        element1.click()

        # Studies List
        element2 = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, studie_expression))
        )
        time.sleep(1)
        element2.click()

        #Protocol List
        element3 = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, protocol_expression))
        )
        time.sleep(1)
        element3.click()

        # Test List
        element4 = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, test_expression))
        )
        element4.click()
        time.sleep(1)

        # Resource List
        element5 = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, resource_expression))
        )
        element5.click()
        time.sleep(1)

        # Your other actions after clicking
        # Allocate Button
        element6 = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, allocate_button))
        ) 
        element6.click()
        time.sleep(1)

        # Your other actions after clicking
         # Connect to the SMTP server and send the email
        with smtplib.SMTP(smtp_server, 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, email_column, message.as_string())

    except Exception as e:
        # Refresh the page
        driver.refresh()
        logging.error(f"Error with some data in row {index+1} ")
        logging.error(f"Error in row {index+1} : {materials},{studies},{protocols},{test},{resources}")

# Second Page Screen 

# Summary & Process Report Button  
summary_button = f'//*[@id="root"]/div/div[1]/div/a/button'

# Download CSV Button 
download_button =f'//*[@id="root"]/div/div/div[2]/div[5]/div[3]/button'

try:
    # Material List
    element7 = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, summary_button))
    )
    element7.click()
    time.sleep(1)

    element8 = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, download_button))
    )
    element8.click()
    time.sleep(1)

    
except Exception as e:
    logging.error(f"Error on the second page: {str(e)}")

email = 'anshika.pal@samta.ai'

# Create a message
message = MIMEMultipart()
message['From'] = sender_email
message['To'] = email
message['Subject'] = 'Summary Report'

# Body of the email
body = f"Hello Sir,\n\nHere are the summary report."

# Attach the body to the message
message.attach(MIMEText(body, 'plain'))

#Define the file to attach
filename = "C:/Users/ashua/Downloads/summary-report.csv"

# Extract just the file name
attachment_filename = os.path.basename(filename)

# Open the file in Python as a binary
attachment = open(filename, 'rb')  # 'r' for read and 'b' for binary

#Encode as base 64
attachment_package = MIMEBase('application', 'octet-stream')
attachment_package.set_payload(attachment.read())
encoders.encode_base64(attachment_package)
attachment_package.add_header('Content-Disposition', f"attachment; filename={attachment_filename}")
message.attach(attachment_package)

# Close the file after attaching it to the email
attachment.close()

# Connect to the SMTP server and send the email
with smtplib.SMTP(smtp_server, 587) as server:
    server.starttls()
    server.login(sender_email, sender_password)
    server.sendmail(sender_email, email, message.as_string())

# Close the browser
driver.quit()

# Get the current date and time
current_datetime = datetime.datetime.now().strftime("%Y-%m-%d_%H_%M_%S")

# Define the new filename with date and time
new_filename = f'summary_report_{current_datetime}.csv'

# Create the full path for the existing and new filenames
old_filepath = "C:/Users/Dell/Downloads/summary-report.csv"
new_filepath = os.path.join("C:/Users/Dell/Downloads/", new_filename)

# Rename the file
os.rename(old_filepath, new_filepath)




py finish
echo `py_result`