import requests
import json
import openpyxl
import concurrent.futures
from collections import OrderedDict

output_file = "location_data.json"

# Performs single location lookup
def un_lookup_single(city, country):

    # Data for the API request
    data ={
                "page": 1,
                "limit": 1,
                "elastic_search": {
                    "location": city,
                    "country_name": country
                },
                "attempt_query": [
                    {
                        "key": "country_name",
                        "operation": "regex",
                        "value": country
                    },
                    {
                        "key": "name",
                        "operation": "regex",
                        "value": city
                    }
                ]
            }

    # Send the API request
    response = requests.post(api_url, headers={"Content-Type": "application/json"}, json=data)

    response_data = {
        "status_code": response.status_code,
        "data": None
    }
    
    if response.ok:
        response_data = response.json()
        return response_data
    else:
        print(f"Request failed: {response.status_code}")
        return None
    
# Function to perform the API requests using multiple threads
def un_lookup(sheet, column1_index, column2_index, stop_at_row=None):
    location_data = [None] * (stop_at_row - 1)

    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = []
        for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if stop_at_row and index > stop_at_row:
                break
            city = row[column1_index]
            country = row[column2_index]
            print(f"Country: {country}\n\t\tCity: {city}")
            
            # Submit the task to the executor
            future = executor.submit(un_lookup_single, city, country)
            futures.append((index -2, future))  # Store row index and future

        # Retrieve results from the executor
        for row_index, future in futures:
            response_data = future.result()
            if response_data:
                location_data[row_index] = response_data

    with open(output_file, "w") as f:
        json.dump(location_data, f, indent=4)

if __name__ == "__main__":
    # API URL for location lookup
    api_url = ""

    excel_file = input("Enter Excel file path: ") # Spreadsheet file path 
    workbook = openpyxl.load_workbook(excel_file)
    sheet_name = "" # Sheet name
    # End row
    stop_at_row = 10
    # Column indices for (city, country)
    column_set = [
        (0, 1)
    ]
    
    sheet = workbook[sheet_name]
    
    # Performs the location lookup for each column set
    for column1_index, column2_index in column_set:
        un_lookup(sheet, column1_index, column2_index, stop_at_row)

    print("Request Fetched!")
