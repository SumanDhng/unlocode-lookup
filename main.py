import requests
import json
import openpyxl
import pandas as pd

output_file = "location_data.json"

# Performs the location lookup
def un_lookup(sheet, column1_index, column2_index,stop_at_row = None):
    location_data = []
    for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True),start=2):
        if stop_at_row and index > stop_at_row:
            break
        city = row[column1_index]
        country = row[column2_index]
        print(f"Country: {country}\n\t\tCity: {city}\t\t")
        
        ## Data for the API request
        # data = {
        #         "page": 1,
        #         "limit": 1,
        #         "elastic_search": {
        #             "location": city,
        #             "country_name": country
        #         },
        #         "attempt_query": [
        #             {
        #                 "key": "country_name",
        #                 "operation": "regex",
        #                 "value": country
        #             },
        #             {
        #                 "key": "name",
        #                 "operation": "regex",
        #                 "value": city
        #             }
        #         ]
        #     }
        
        # Data for the revised API request
        data = {
                    "page": 1,
                    "limit": 1,
                    "elastic_search": f"{country}, {city}",
                    "query": [
                        {
                            "key": "country",
                            "operation": "regex",
                            "value": country
                        },
                        {
                            "key": "location",
                            "operation": "regex",
                            "value": city
                        }
                    ]
                }

        response = requests.post(api_url, headers={"Content-Type": "application/json"},json = data)

        if response.ok:
            response_data = response.json()
            location_data.append(response_data)
            with open(output_file, "w") as f:
                json.dump(location_data, f, indent=4)
            

        else:
            print(f"Request failed: {response.status_code}")


if __name__ == "__main__":
    
    # API URL for location lookup
    # api_url = ""
    
    # Revised API URL for location lookup
    api_url = "https://appserver.qa.rippey.ai/operations/dataDictionaries/v2/manifest/city_port_revised/dictionary/query"

    # Spreadsheet file path 
    excel_file = input("Enter Excel file path: ")
    workbook = openpyxl.load_workbook(excel_file)
    # Sheet name
    sheet_name = ""
    # Stopping row number
    stop_at_row = 10
    # Set order (city, country)
    column_set = [
         (0,1)
    ]
    
    sheet = workbook[sheet_name]
    
    # Performs the location lookup for each column set
    for column1_index, column2_index in column_set:
        un_lookup(sheet, column1_index, column2_index,stop_at_row)
    
    
    print("DONE!")