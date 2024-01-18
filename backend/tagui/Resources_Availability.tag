py begin

import requests
import pandas as pd
import logging

# Configure the logging module
log_path = 'D:\Edward\TagUI\Mapped_Data.xlsx'
logging.basicConfig(filename=log_path, level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

def extract_data():
    url = 'https://6596915d6bb4ec36ca02eba3.mockapi.io/resource'

    excel_path = 'D:\Edward\TagUI\Mapped_Data.xlsx'
    df = pd.read_excel(excel_path, sheet_name="Sheet1")

    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an exception for error status codes

        all_data = response.json()  # Assuming the response is in JSON format

        for data in all_data:
            if data.get("status") == "Available" and data.get("is_internal") == "True":
                for index, row in df.iterrows():
                    if row['test Skills'] == data.get("skills"):
                        df.at[index, "resources"] = data.get("name")
                        df.at[index, "emp_id"] = data.get("emp_id")
                        data["status"] = "Not Available"
                        break
            elif data.get("status") == "Available":
                for index, row in df.iterrows():
                    if row['test skills'] == data.get("skills") and pd.isna(row["resources"]):
                        df.at[index, "resources"] = data.get("name")
                        df.at[index, "emp_id"] = data.get("emp_id")
                        data["status"] = "Not Available"
                        break

        for data in all_data:
            if data.get("status") == "Available" and data.get("is_internal") == "False":
                for index, row in df.iterrows():
                    if row['Test Skills'] == data.get("skills") and pd.isna(row["resources"]):
                        df.at[index, "resources"] = data.get("name")
                        df.at[index, "emp_id"] = data.get("emp_id")
                        data["status"] = "Not Available"
                        break

        columns_to_replace = ['resources']
        df[columns_to_replace] = df[columns_to_replace].fillna("No Resource Available")
        df.to_excel(excel_path, sheet_name="Sheet1", index=False)
        return []

    except Exception as e:
        logging.error("Error fetching data: %s", e)
        return [str(e)]

# Call the function
errors = extract_data()

if errors:
    print("Errors occurred. Check the log file for details:", log_path)
else:
    print("Data extraction completed successfully.")



py finish
echo `py_result`
