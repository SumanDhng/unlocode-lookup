# Location Data Extraction

This script extracts location data using an API for a given set of cities and countries listed in an Excel file. It performs an API request for each city-country pair and saves the response data in a JSON file.

## Requirements

- `requests`, `openpyxl`, and `pandas` libraries

- Install the required dependencies using the following command:

    ```
    pip install -r requirements.txt
    ```


## Installation

1. Clone the repository or download the script.

2. Provide API URL, number of Rows and Sheet name to process

3. Run the script by executing the following command:

    ```
    py main.py
    ```

4. Follow the prompts to provide the necessary inputs, including the Excel file path and the API URL for location data extraction.

## Usage

1. Make sure the Excel file contains the city and country information in the specified format.
2. Check the script output for any error messages or status codes returned during the API requests.
3. Review the generated JSON file for the extracted location data.

## Note

Ensure that the API endpoint is functional and accessible from the script's environment.