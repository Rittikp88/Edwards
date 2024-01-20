py begin

import requests
import pandas as pd
import logging

# Configure logging
logging.basicConfig(filename='D:\Edward\TagUI\Mapped_data_error_log.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s: %(message)s')

def extract_data():
    
    url = 'https://6596915d6bb4ec36ca02eba3.mockapi.io/resource'

    excel_path = 'D:\Edward\TagUI\Mapped_Data.xlsx'
    df=pd.read_excel(excel_path, sheet_name="Sheet1")
    df["resources"] = None
    df["status"] = None
    df["emp_id"] = None

    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an exception for error status codes

        all_data = response.json()  # Assuming the response is in JSON format

        response= 'Success'
        for data in all_data:
            if data.get("status")=="Available" and data.get("is_internal")==True:
                for index, row in df.iterrows():
                    if row['test skills']==data.get("skills"):
                        df.at[index, "resources"] = data.get("name")
                        df.at[index, "emp_id"] = data.get("emp_id")
                        df.at[index, "status"] = "Allocated"
                        data["status"]="Not Available"
                        #print(df, "after first true condition")
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
            if data.get("status")=="Available" and data.get("is_internal")==False:
                for index, row in df.iterrows():
                    if row['test skills']==data.get("skills") and pd.isna(row["resources"]):
                        df.at[index, "resources"] = data.get("name")
                        df.at[index, "emp_id"] = data.get("emp_id")
                        df.at[index, "status"] = "Allocated"
                        data["status"]="Not Available"
                        break
        #print(df,"Df value")
        columns_to_replace = ['resources']
        df[columns_to_replace] = df[columns_to_replace].fillna("No Resource Available")
        columns_to_replace = ['status']
        df[columns_to_replace] = df[columns_to_replace].fillna("Null")
        columns_to_replace = ['emp_id']
        df[columns_to_replace] = df[columns_to_replace].fillna("Null")
        #print(df,"Df record")
        df.to_excel(excel_path, sheet_name="Sheet1", index=False)
        return response
    
    except Exception as e:
        logging.exception(e)
        print("Error fetching data:", e)

extract_data()



py finish
echo `py_result`
