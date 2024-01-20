py begin

from datetime import datetime, timedelta
import pandas as pd
import requests
import json
import logging
import random
import numpy as np

# Configure logging
logging.basicConfig(filename='D:\Edward\TagUI\Mapped_Data_error.log', level=logging.ERROR, format='%(asctime)s - %(levelname)s: %(message)s')

def test_allocation():
    # Step 1: Read the Excel file with two sheets
    excel_file_path = 'D:\Edward\TagUI\Mapped_Data.xlsx'
    excel_data_sheets = pd.read_excel(excel_file_path, sheet_name='Sheet1')
    excel_data_sheets = excel_data_sheets.replace({np.nan:None})
    json_data = excel_data_sheets.to_dict(orient='records')
    # print(json_data)

    # Step 2: Convert Excel data to Python dictionaries and create or update resources
    api_base_url = 'https://6596915d6bb4ec36ca02eba3.mockapi.io/resource'
    response = requests.get(api_base_url)
    response_data = response.json()
    existing_emp_ids = set(record['emp_id'] for record in response_data)
    existing_emp_ids=list(existing_emp_ids)

    # List to store records
    new_records = []

    cureent_date = datetime.now().date()
        
    material_study_protocols = {}

    try:
        # print(json_data,"json_data")
        for record in json_data:
            # print(record,"record")
            material = record['Materials']
            study = record['Studies']
            protocol = record['Protocols']
            duration = record['duration']
            test = record['test']
            resource = record['resources']
            no_resource_message = "No Resource Available"

            
            if no_resource_message == resource or resource is None:
                logging.error(f"no resource available in row: {json_data.index(record)+1}")
                logging.error(f"Error in row {json_data.index(record)+1} : {material},{study},{protocol},{test}")
                continue
            key_name = material + '_' + study + '_' + protocol

            start_date = material_study_protocols.setdefault(key_name, datetime.now().date())
            end_date = start_date + timedelta(days=duration-1)
        
            material_study_protocols[key_name]= end_date
        
            record['startDate'] = start_date
            record['endDate'] = end_date

            if record['emp_id'] in existing_emp_ids:
                # print(response_data,"response_data")
                for data in response_data:
                    if data['emp_id']==record['emp_id']:
                        id=data['id']
                        record['EQL'] = record['emp_id'] + str(random.randint(0,9))
                        api_endpoint = f'{api_base_url}/{id}'
                        response = requests.put(api_endpoint, data=record)

                        print(response)
    except Exception as e:
        logging.exception(e)
        print(e)           

test_allocation()


py finish
echo `py_result`