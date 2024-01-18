py begin

from datetime import datetime, timedelta
import pandas as pd
import requests
import random
import logging

# Configure logging
logging.basicConfig(filename='error_log.txt', level=logging.ERROR, format='%(asctime)s - %(levelname)s: %(message)s')

format = "%Y-%m-%d"

def test_allocation():
    # Step 1: Read the Excel file with two sheets
    excel_file_path = 'D:\Edward\TagUI\Mapped_Data.xlsx'
    excel_data_sheets = pd.read_excel(excel_file_path, sheet_name='Sheet1')
    json_data = excel_data_sheets.to_dict(orient='records')

    # Step 2: Convert Excel data to Python dictionaries and create or update resources
    api_base_url = 'https://6596915d6bb4ec36ca02eba3.mockapi.io/resource'
    response = requests.get(api_base_url)
    response_data = response.json()
    existing_emp_ids = set(record['emp_id'] for record in response_data)

    try:
        for record in json_data:
            material, study, protocol, duration = record['Materials'], record['Studies'], record['Protocols'], record['duration']
            key_name = f'{material}_{study}_{protocol}'

            start_date = datetime.now() if key_name not in material_study_protocols else material_study_protocols[key_name]
            end_date = start_date + timedelta(days=duration - 1)
            material_study_protocols[key_name] = end_date

            record.update({'startDate': start_date.strftime(format), 'endDate': end_date.strftime(format), 'EQL': f"{record['emp_id']}{random.randint(0, 9)}"})

            if record['emp_id'] in existing_emp_ids:
                id = next((data['id'] for data in response_data if data['emp_id'] == record['emp_id']), None)
                if id:
                    api_endpoint = f'{api_base_url}/{id}'
                    response = requests.put(api_endpoint, data=record)
                    print(response)
    except Exception as e:
        logging.error(f'Error: {e}', exc_info=True)

# Dictionary to store start and end dates for material, study, protocol combination
material_study_protocols = {}

test_allocation()


py finish
echo `py_result`