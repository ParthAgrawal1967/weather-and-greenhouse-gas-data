import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
import time
import os

# Function to fetch greenhouse gas data from WAQI API
def get_gas_data(api_key, city):
    # base_url = "/https://api.waqi.info/feed/"
    base_url = f'https://aqicn.org/city/india/lucknow/talkatora'
    print(base_url)
    params = {"token": api_key, "city": city}
    print(f'base_url+token:{api_key}+city:{city}')
    # response = requests.get(base_url + city, params=params)
    response = requests.get('https://aqicn.org/city/india/lucknow/talkatora')
    print(response.text)
    data = response.json()
    print(data)
    return data

# Function to update Excel file with greenhouse gas data
def update_gas_table(gas_data, excel_writer):
    # Check if the 'GreenhouseGas' sheet already exists in the Excel file
    if 'GreenhouseGas' in excel_writer.book.sheetnames:
        # Load existing greenhouse gas table
        gas_df = pd.read_excel(excel_writer, sheet_name="GreenhouseGas")

        # Check if the 'data' key exists in the response
        if "data" in gas_data:
            # Extract relevant information from the response
            timestamp = datetime.utcfromtimestamp(gas_data["data"]["time"]["s"]).strftime('%Y-%m-%d %H:%M:%S')
            gas_level = gas_data["data"]["iaqi"].get("pm25", {}).get("v", None)

            # Check if gas_level is available
            if gas_level is not None:
                # Create a new DataFrame with the extracted data
                new_data = pd.DataFrame({
                    "Timestamp": [timestamp],
                    "Gas Level": [gas_level],
                    # Add more parameters if needed
                })

                # Check if the last entry in the table has the same data as the new data
                if gas_df.empty or not last_entry_equals_new_data(gas_df, new_data):
                    # Append new data to the existing DataFrame if it's different
                    gas_df = pd.concat([gas_df, new_data], ignore_index=True)

                    # Write the updated DataFrame to the 'GreenhouseGas' sheet
                    gas_df.to_excel(excel_writer, sheet_name="GreenhouseGas", index=False, header=True)

# Function to check if the last entry in the table is equal to the new data
def last_entry_equals_new_data(dataframe, new_data):
    last_entry = dataframe.iloc[-1, :]
    return last_entry.to_dict() == new_data.iloc[0, :].to_dict()

# WAQI API key
waqi_api_key = "YOUR_WAQI_API_KEY"

# City and Excel file setup
city = "Lucknow"
excel_file = "greenhouse_gas_data_waqi.xlsx"

while True:
    # Check if the Excel file exists
    file_exists = os.path.isfile(excel_file)

    # Create or load Excel file
    with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        try:
            # Update greenhouse gas table with new data
            gas_data = get_gas_data(waqi_api_key, city)
            update_gas_table(gas_data, writer)

        except FileNotFoundError:
            print(f"Error: File '{excel_file}' not found.")

    time.sleep(1)  # Update every second for the greenhouse gas table
