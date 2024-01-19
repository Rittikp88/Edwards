py begin

import requests
import pandas as pd
import logging

def extract_data():

	# Configure logging to write to a file
    log_file_path = 'D:\Edward\TagUI\extract_data_log.txt'
    logging.basicConfig(filename=log_file_path, level=logging.ERROR)
	
    url = 'https://6596915d6bb4ec36ca02eba3.mockapi.io/resource'

    excel_path = 'D:\Edward\TagUI\Mapped_Data.xlsx'
    df=pd.read_excel(excel_path, sheet_name="Sheet1")

    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an exception for error status codes

        all_data = response.json()  # Assuming the response is in JSON format

        response=[]
        for data in all_data:
            response_dict={}
            if data.get("status")=="Available" and data.get("is_internal")=="True":
                for index, row in df.iterrows():
                    if row['test Skills']==data.get("skills"):
                        df.at[index, "resources"] = data.get("name")
                        df.at[index, "emp_id"] = data.get("emp_id")
                        df.at[index, "status"] = "Allocated"
                        data["status"]="Not Available"
                        break
            elif data.get("status")=="Available":
                for index, row in df.iterrows():
                    if row['test skills']==data.get("skills") and pd.isna(row["resources"]):
                        df.at[index, "resources"] = data.get("name")
                        df.at[index, "emp_id"] = data.get("emp_id")
                        df.at[index, "status"] = "Allocated"
                        data["status"]="Not Available"
                        break

        for data in all_data:
            response_dict={}
            if data.get("status")=="Available" and data.get("is_internal")=="False":
                for index, row in df.iterrows():
                    if row['Test Skills']==data.get("skills") and pd.isna(row["resources"]):
                        df.at[index, "resources"] = data.get("name")
                        df.at[index, "emp_id"] = data.get("emp_id")
                        df.at[index, "status"] = "Allocated"
                        data["status"]="Not Available"
                        break
        columns_to_replace = ['resources']
        df[columns_to_replace] = df[columns_to_replace].fillna("No Resource Available")
        columns_to_replace = ['status']
        df[columns_to_replace] = df[columns_to_replace].fillna("Null")
        columns_to_replace = ['emp_id']
        df[columns_to_replace] = df[columns_to_replace].fillna("Null")
        df.to_excel(excel_path, sheet_name="Sheet1", index=False)
        return response
    
    except Exception as e:
		# Log the error
        logging.error("Error fetching data: %s", e)
        #print("Error fetching data:", e)

extract_data()



py finish
echo `py_result`
