import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
import time
import os

# Function to fetch weather data from OpenWeatherMap API
def get_weather_data(api_key, city):
    base_url = "http://api.openweathermap.org/data/2.5/weather"
    params = {"q": city, "appid": api_key}
    response = requests.get(base_url, params=params)
    data = response.json()
    return data

# Function to update Excel file with weather data
def update_weather_table(weather_data, excel_writer):
    # Check if the 'Weather' sheet already exists in the Excel file
    if 'Weather' in excel_writer.book.sheetnames:
        # Load existing weather table
        weather_df = pd.read_excel(excel_writer, sheet_name="Weather")
    else:
        # Create new DataFrame if the sheet doesn't exist
        weather_df = pd.DataFrame(columns=["Time", "Temperature", "Humidity", "Description", "Wind Speed", 
                                           "Pressure", "Visibility", "Cloudiness", "Sunrise", "Sunset"])

    # Append new data to the existing DataFrame
    new_data = pd.DataFrame({
        "Time": [start_date],  # Use the start_date for the initial time
        "Temperature": [weather_data["main"]["temp"]],
        "Humidity": [weather_data["main"]["humidity"]],
        "Description": [weather_data["weather"][0]["description"]],
        "Wind Speed": [weather_data["wind"]["speed"]],
        "Pressure": [weather_data["main"]["pressure"]],
        "Visibility": [weather_data["visibility"]],
        "Cloudiness": [weather_data["clouds"]["all"]],
        "Sunrise": [datetime.fromtimestamp(weather_data["sys"]["sunrise"], timezone.utc)],
        "Sunset": [datetime.fromtimestamp(weather_data["sys"]["sunset"], timezone.utc)],
    })

    # Convert datetime values to string representation
    new_data["Sunrise"] = new_data["Sunrise"].dt.strftime('%Y-%m-%d %H:%M:%S')
    new_data["Sunset"] = new_data["Sunset"].dt.strftime('%Y-%m-%d %H:%M:%S')

    # Append new data to the existing DataFrame
    weather_df = pd.concat([weather_df, new_data], ignore_index=True)

    # Write the updated DataFrame to the 'Weather' sheet
    weather_df.to_excel(excel_writer, sheet_name="Weather", index=False, header=True)

# OpenWeatherMap API key
weather_api_key = "80d6139c7432cfe952765d3fb08ada00"

# City and Excel file setup
city = "Lucknow"
excel_file = "weather_data.xlsx"

# Calculate the start date as 3 years ago from the current date
start_date = datetime.now() - timedelta(days=3*365)

while True:
    # Check if the Excel file exists
    file_exists = os.path.isfile(excel_file)

    # Create or load Excel file
    with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        try:
            # Update weather table with new data
            weather_data = get_weather_data(weather_api_key, city)
            update_weather_table(weather_data, writer)

        except FileNotFoundError:
            print(f"Error: File '{excel_file}' not found.")

    time.sleep(1)  # Update every second for the weather table

