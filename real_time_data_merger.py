import pandas as pd
import time

def merge_excel_files(file1, file2, output_file):
    while True:
        try:
            # Read data from source files
            df_weather = pd.read_excel(file1)
            df_gas = pd.read_excel(file2)

            # Concatenate data from both files (union of unique rows)
            merged_df = pd.concat([df_weather, df_gas]).drop_duplicates().reset_index(drop=True)

            # Write merged data to the output file
            merged_df.to_excel(output_file, index=False)

            print("Merged data successfully. Waiting for the next update...")

            # Sleep for one second before checking for updates again
            time.sleep(1)

        except Exception as e:
            print(f"Error: {str(e)}")

if __name__ == "__main__":
    # Specify the file paths
    file1_path = r"C:\Weather ML Project\weather_data.xlsx"
    file2_path = r"C:\Weather ML Project\greenhouse_gas_data.xlsx"
    output_file_path = r"C:\Weather ML Project\merged_data.xlsx"

    # Call the merge function
    merge_excel_files(file1_path, file2_path, output_file_path)