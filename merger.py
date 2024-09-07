import pandas as pd
import os
import sys

# Check if the folder path is provided as an argument
if len(sys.argv) < 2:
    print("Please provide the folder name containing the CSV files as an argument.")
    sys.exit(1)

# Get the folder path from command-line argument
input_folder = sys.argv[1]

# Check if the folder exists
if not os.path.exists(input_folder):
    print(f"The folder {input_folder} does not exist.")
    sys.exit(1)

# Define the output file name
output_file = "merged_output.csv"

# Get all CSV files in the folder
csv_files = [f for f in os.listdir(input_folder) if f.endswith('.csv')]

# Initialize an empty DataFrame to store the merged data
merged_data = pd.DataFrame()

# Loop through each file
for csv_file in csv_files:
    file_path = os.path.join(input_folder, csv_file)
    
    # Read the CSV file, skipping the first 13 rows and using the 14th row as header
    df = pd.read_csv(file_path, skiprows=13)
    
    # Append the data to the merged dataframe
    if merged_data.empty:
        # If it's the first file, keep the header
        merged_data = df
    else:
        # For subsequent files, append the data without headers
        merged_data = pd.concat([merged_data, df], ignore_index=True)

# Save the merged data to a new CSV file
merged_data.to_csv(output_file, index=False)

print(f"CSV files merged successfully into {output_file}")
