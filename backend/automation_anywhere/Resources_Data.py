import pandas as pd
import requests
import json
import logging

# Configure logging
logging.basicConfig(filename='error_log.txt', level=logging.ERROR, format='%(asctime)s - %(levelname)s: %(message)s')

def create_and_update_resources():
    # Step 1: Read the Excel file with two sheets
    excel_file_path = 'D:/Edward/Available Resource Data.xlsx'
    excel_data_sheets = pd.read_excel(excel_file_path, sheet_name='Sheet1')
    json_data = excel_data_sheets.to_dict(orient='records')

    # Step 2: Convert Excel data to Python dictionaries and create or update resources
    api_base_url = 'https://6596915d6bb4ec36ca02eba3.mockapi.io/resource'
    response = requests.get(api_base_url)
    response_data = response.json()

    # Create a set of existing employee IDs for faster lookup
    existing_emp_ids = set(record['emp_id'] for record in response_data)

    # List to store records
    new_records = []

    try:
        for record in json_data:
            if record['emp_id'] not in existing_emp_ids:
                print("Not matched, creating resource:", record)
                response = requests.post(api_base_url, json=record)
                new_records.append(record)
            else:
                print("Matched, skipping resource:", record)
                api_endpoint = f'{api_base_url}/{id}'
                response = requests.put(api_base_url, json=record)
        
    except Exception as e:
        error_message = f"Error processing record. Error: {str(e)}"
        print(error_message)
        logging.error(error_message)

    # Dump new records to a JSON file
    with open('new_records.json', 'w') as json_file:
        json.dump(new_records, json_file, indent=2)

    return "Resource creation/update process completed."

# Call the function
result = create_and_update_resources()
print(result)
