py begin

import pandas as pd
import requests
import json
import logging

# Configure logging
logging.basicConfig(filename='error_log.txt', level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

def update_data():
    # Step 1: Read the Excel file with two sheets
    excel_file_path = 'D:\Edward\TagUI\Leave Update.xlsx'

    try:
        excel_data_sheets = pd.read_excel(excel_file_path, sheet_name='Sheet1')
    except Exception as e:
        logging.error(f"Error reading Excel file: {str(e)}")
        return "Update process failed."

    # Step 2: Convert Excel data to Python dictionaries and update the data
    api_base_url = 'https://6596915d6bb4ec36ca02eba3.mockapi.io/resource'

    for sheet_name, sheet_data in excel_data_sheets.iterrows():
        record_id = sheet_data['id']
        logging.info(f"Processing record with ID: {record_id}")

        try:
            if record_id:
                # Modify the URL to include the record ID
                api_endpoint = f'{api_base_url}/{record_id}'

                # Remove the 'id' field from the record before sending the request
                record = {
                    "id": sheet_data['id'],
                    "status": sheet_data['status'],
                    "duration": sheet_data['duration'],
					"startDate": sheet_data['startDate'],
                    "endDate": sheet_data['endDate']
                }
                # Make the PUT request
                put_data = record
                
                response = requests.put(api_endpoint, data=record)
                # Check the response status
                logging.info(response.text)
                if response.status_code == 200:
                    logging.info("Data updated Successfully")
                else:
                    logging.error(f"Response status: {response.status_code}")
            else:
                logging.warning("Skipping record without ID in sheet")
        except Exception as e:
            logging.error(f"Error processing record with ID {record_id}: {str(e)}")

    return "Update process completed."

# Call the function

result = update_data()
print(result)


py finish
echo `py_result`
