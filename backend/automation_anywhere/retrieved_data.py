import requests
import pandas as pd

def extract_data():
    url = 'https://6596915d6bb4ec36ca02eba3.mockapi.io/resource'
    excel_path = 'D:/Edward/RetrievedData.xlsx'

    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an exception for error status codes
        all_data = response.json()  # Assuming the response is in JSON format

        data_list = []
        for data in all_data:
            # Extracting fields
            data_dict = {
                "id": data.get("id"),
                "name": data.get("name"),
                "status": data.get("status"),
                "skills": data.get("skills"),
                "test": data.get("test"),
                "is_internal": data.get("is_internal"),
                "emp_id": data.get("emp_id"),
                "email id": data.get("email id"),
                "Materials": data.get("Materials"),
                "Studies": data.get("Studies"),
                "Protocols": data.get("Protocols"),
                "duration": data.get("duration"),
                "test skills": data.get("test skills"),
                "startDate": data.get("startDate"),
                "endDate": data.get("endDate"),
                "EQL": data.get("EQL")
              
            }
            data_list.append(data_dict)
            # print(data_dict)
        print(data_list)

        # Creating DataFrame
        new_df = pd.DataFrame(data_list)
    

        # Read existing Excel file if it exists
        try:
            existing_df = pd.read_excel(excel_path, sheet_name="Sheet1")
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        except FileNotFoundError:
            combined_df = new_df

        # Save DataFrame to Excel
        combined_df.to_excel(excel_path, sheet_name="Sheet1", index=False)

    except Exception as e:
        print("Error fetching or processing data:", e)

    return f"Done"

# Call the function
extract_data()
