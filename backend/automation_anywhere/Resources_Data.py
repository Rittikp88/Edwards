import pandas as pd
import requests
import json
import logging

# Configure logging
logging.basicConfig(filename='error_log.txt', level=logging.ERROR, format='%(asctime)s - %(levelname)s: %(message)s')

def create_and_dump_to_json():
    # Step 1: Read the Excel file with two sheets
    excel_file_path = 'D:/Edward/Available Resource Data.xlsx'
    excel_data_sheets = pd.read_excel(excel_file_path, sheet_name='Sheet1')

    # Step 2: Convert Excel data to Python dictionaries and create new resources
    api_base_url = 'https://6596915d6bb4ec36ca02eba3.mockapi.io/resource'

    # List to store records
    records = []

    for _, sheet_data in excel_data_sheets.iterrows():
        try:
            # Modify the URL to include the record ID
            api_endpoint = api_base_url

            # Remove the 'id' field from the record before sending the request
            record = {
                "name": sheet_data['name'],
                "status": sheet_data['status'],
                "skills": sheet_data['skills'],
                "test": sheet_data['test'],
                "is_internal": sheet_data['is_internal']
            }

            print(record)

            # Make the POST request
            response = requests.post(api_endpoint, json=record)

            # Check the response status
            print(response.text)
            if response.status_code == 201:  # Assuming 201 Created is returned for successful resource creation
                print(f"Data created Successfully")
                records.append(record)
            else:
                error_message = f"Error creating data. Response status: {response.status_code}, Response text: {response.text}"
                print(error_message)
                logging.error(error_message)
        except Exception as e:
            error_message = f"Error processing record. Error: {str(e)}"
            print(error_message)
            logging.error(error_message)

    # Dump records to a JSON file
    with open('output.json', 'w') as json_file:
        json.dump(records, json_file, indent=2)

    # Fetch all details from the database
    try:
        # Make the GET request to retrieve all records
        response = requests.get(api_base_url)

        # Check the response status
        if response.status_code == 200:  # Assuming 200 OK is returned for successful retrieval
            # Parse the JSON response
            db_records = response.json()
            print(f"Data retrieved Successfully")

            # Convert the retrieved data to a DataFrame
            db_df = pd.DataFrame(db_records)

            # Modify the path to the new Excel file
            new_excel_file_path = 'D:/Edward/RetrievedData.xlsx'

            # Write the DataFrame to a new Excel sheet at a different address
            with pd.ExcelWriter(new_excel_file_path, engine='openpyxl', mode='w') as writer:
                db_df.to_excel(writer, sheet_name='RetrievedData', index=False)

        else:
            error_message = f"Error retrieving data. Response status: {response.status_code}, Response text: {response.text}"
            print(error_message)
            logging.error(error_message)
    except Exception as e:
        error_message = f"Error retrieving data. Error: {str(e)}"
        print(error_message)
        logging.error(error_message)

    return "Creation and JSON dump process completed."

# Call the function
result = create_and_dump_to_json()
print(result)
