py begin

import pandas as pd
import requests
import json
import logging

# Configure logging
logging.basicConfig(filename='D:\Edward\TagUI\Available_resource_data_error.log', level=logging.ERROR, format='%(asctime)s - %(levelname)s: %(message)s')

def create_update_and_fetch_data():
    # Step 1: Read the Excel file with two sheets
    excel_file_path = 'D:\Edward\TagUI\Available Resource Data.xlsx'
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
                matched_record = next(item for item in response_data if item['emp_id'] == record['emp_id'])
                print("Matched, updating resource:", record)
                api_endpoint = f'{api_base_url}/{matched_record["id"]}'
                response = requests.put(api_endpoint, json=record)
        
    except Exception as e:
        error_message = f"Error processing record. Error: {str(e)}"
        print(error_message)
        logging.error(error_message)

    # Dump new records to a JSON file
    with open('new_records.json', 'w') as json_file:
        json.dump(new_records, json_file, indent=2)

    # Step 3: Fetch data from the API and save it to Excel file
    api_base_url = 'https://6596915d6bb4ec36ca02eba3.mockapi.io/resource'
    response = requests.get(api_base_url)
    response_data = response.json()

    # Convert API response to DataFrame
    df = pd.DataFrame(response_data)

    # Save DataFrame to Excel file
    retrieved_excel_file_path = 'D:\Edward\TagUI\RetrievedData.xlsx'
    df.to_excel(retrieved_excel_file_path, index=False)

    return "Resource creation/update and data retrieval process completed."

# Call the function
result = create_update_and_fetch_data()
print(result)



py finish
echo `py_result`
