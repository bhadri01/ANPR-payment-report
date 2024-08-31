import pandas as pd
from datetime import datetime

def process_csv(input_file):
    # Read the data from the provided CSV file, skipping the first 14 rows
    data = pd.read_csv(input_file, skiprows=13)

    # Convert 'Payment Date' to datetime and normalize to just the date (remove time)
    data['Payment Date'] = pd.to_datetime(data['Payment Date'], errors='coerce').dt.date

    # Handle cases where dates couldn't be converted (if any)
    if data['Payment Date'].isnull().any():
        print("Some dates could not be converted and will be ignored.")

    # Filter out rows where 'Payment Date' couldn't be parsed
    data = data.dropna(subset=['Payment Date'])

    # Create a date range from 01.12.2020 to today, normalizing to date only
    start_date = datetime(2020, 12, 1).date()
    end_date = datetime.today().date()
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    # Convert date_range to datetime.date
    date_range = [d.date() for d in date_range]

    # Initialize the processed data DataFrame
    processed_data = pd.DataFrame({
        "Date": date_range,
        "No.of Cases Fine Collected": 0,
        "Total No. of 100 Rs Cases": 0,
        "Collected Fine Amount in 100 Rs": 0,
        "Total No. of 200 Rs Cases": 0,  # Added for 200 Rs cases
        "Collected Fine Amount in 200 Rs": 0,  # Added for 200 Rs cases
        "Total No. of 1000 Rs Cases": 0,
        "Collected Fine Amount in 1000 Rs": 0,
        "Total No. of Cases Fine Collected": 0,
        "Total Fine Amount Collected": 0
    })

    # Process data
    for index, row in data.iterrows():
        date = row['Payment Date']
        challan_amount = row['Challan Amount']

        if date in processed_data['Date'].values:
            date_row = processed_data[processed_data['Date'] == date].index[0]

            if challan_amount == 100:
                processed_data.at[date_row, 'Total No. of 100 Rs Cases'] += 1
                processed_data.at[date_row, 'Collected Fine Amount in 100 Rs'] += challan_amount
            elif challan_amount == 200:
                processed_data.at[date_row, 'Total No. of 200 Rs Cases'] += 1
                processed_data.at[date_row, 'Collected Fine Amount in 200 Rs'] += challan_amount
            elif challan_amount == 1000:
                processed_data.at[date_row, 'Total No. of 1000 Rs Cases'] += 1
                processed_data.at[date_row, 'Collected Fine Amount in 1000 Rs'] += challan_amount

            processed_data.at[date_row, 'Total No. of Cases Fine Collected'] += 1
            processed_data.at[date_row, 'Total Fine Amount Collected'] += challan_amount
            processed_data.at[date_row, 'No.of Cases Fine Collected'] += 1

    # Save the new DataFrame to Excel
    processed_data.to_excel('Processed_ANPR_Fine_Details.xlsx', index=False)

# Example usage:
process_csv('ANPR/2.Jan_2021.csv')
