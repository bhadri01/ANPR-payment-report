import pandas as pd
from datetime import datetime
import os

def process_all_csvs(directory, generate_daily, generate_monthly, start_date, end_date):
    # Initialize the processed data DataFrame with the full date range
    start_date_obj = datetime(2020, 12, 1).date()
    end_date_obj = datetime.today().date()
    date_range = pd.date_range(start=start_date_obj, end=end_date_obj, freq='D')
    date_range = [d.date() for d in date_range]  # Convert to datetime.date

    # Create the output DataFrame
    final_processed_data = pd.DataFrame({
        "Date": date_range,
        "No.of Cases Fine Collected": 0,
        "Total No. of 100 Rs Cases": 0,
        "Collected Fine Amount in 100 Rs": 0,
        "Total No. of 200 Rs Cases": 0,
        "Collected Fine Amount in 200 Rs": 0,
        "Total No. of 1000 Rs Cases": 0,
        "Collected Fine Amount in 1000 Rs": 0,
        "Total No. of Cases Fine Collected": 0,
        "Total Fine Amount Collected": 0
    })

    # Iterate over each file in the directory
    for filename in os.listdir(directory):
        if filename.endswith('.csv'):
            filepath = os.path.join(directory, filename)
            print(f"Processing file: {filename}")

            # Try reading the CSV with varying numbers of rows skipped
            header_found = False
            for skip in range(0, 21):
                try:
                    data = pd.read_csv(filepath, skiprows=skip)
                    
                    # Check if the required columns exist
                    if 'Payment Date' in data.columns and 'Challan Amount' in data.columns:
                        header_found = True
                        break
                except pd.errors.ParserError:
                    continue

            if not header_found:
                # If no valid header was found, skip this file
                print(f"Valid header not found in {filename}. Skipping this file.")
                continue

            # Convert 'Payment Date' to datetime and normalize to just the date
            data['Payment Date'] = pd.to_datetime(data['Payment Date'], errors='coerce').dt.date

            # Handle cases where dates couldn't be converted (if any)
            if data['Payment Date'].isnull().any():
                print(f"Some dates in {filename} could not be converted and will be ignored.")

            # Filter out rows where 'Payment Date' couldn't be parsed
            data = data.dropna(subset=['Payment Date'])

            # Process data
            for _, row in data.iterrows():
                date = row['Payment Date']
                challan_amount = row['Challan Amount']

                if date in final_processed_data['Date'].values:
                    date_row = final_processed_data[final_processed_data['Date'] == date].index[0]

                    if challan_amount == 100:
                        final_processed_data.at[date_row, 'Total No. of 100 Rs Cases'] += 1
                        final_processed_data.at[date_row, 'Collected Fine Amount in 100 Rs'] += challan_amount
                    elif challan_amount == 200:
                        final_processed_data.at[date_row, 'Total No. of 200 Rs Cases'] += 1
                        final_processed_data.at[date_row, 'Collected Fine Amount in 200 Rs'] += challan_amount
                    elif challan_amount == 1000:
                        final_processed_data.at[date_row, 'Total No. of 1000 Rs Cases'] += 1
                        final_processed_data.at[date_row, 'Collected Fine Amount in 1000 Rs'] += challan_amount

                    final_processed_data.at[date_row, 'Total No. of Cases Fine Collected'] += 1
                    final_processed_data.at[date_row, 'Total Fine Amount Collected'] += challan_amount
                    final_processed_data.at[date_row, 'No.of Cases Fine Collected'] += 1

            print(f"Done processing file: {filename}")

    # Calculate and append the sum row for the daily data
    if generate_daily:
        sum_row_daily = final_processed_data.sum(numeric_only=True)
        sum_row_daily['Date'] = 'Total'
        final_processed_data = final_processed_data._append(sum_row_daily, ignore_index=True)

        # Save the final aggregated daily DataFrame to Excel
        final_processed_data.to_excel('final_details_daily.xlsx', index=False)
        print("Daily details saved to 'final_details_daily.xlsx'.")

    # Define numeric columns for use in both monthly and custom reports
    numeric_cols = final_processed_data.columns.drop('Date')

    # Create a month-wise summary by grouping and summing the daily data
    if generate_monthly:
        final_processed_data['Month'] = pd.to_datetime(final_processed_data['Date'], errors='coerce').dt.to_period('M')
        
        # Select numeric columns for summation, excluding the 'Date' column
        numeric_cols_monthly = final_processed_data.columns.drop(['Date', 'Month'])
        
        # Group by 'Month' and sum the numeric columns
        month_wise_summary = final_processed_data.groupby('Month')[numeric_cols_monthly].sum().reset_index()

        # Calculate and append the sum row for the monthly data
        sum_row_monthly = month_wise_summary.sum(numeric_only=True)
        sum_row_monthly['Month'] = 'Total'
        month_wise_summary = month_wise_summary._append(sum_row_monthly, ignore_index=True)

        # Save the month-wise summary to Excel
        month_wise_summary.to_excel('final_details_monthly.xlsx', index=False)
        print("Monthly summary saved to 'final_details_monthly.xlsx'.")

    # Generate a report for the specified date range
    if start_date and end_date:
        filtered_data = final_processed_data[
            (final_processed_data['Date'] >= start_date) &
            (final_processed_data['Date'] <= end_date)
        ]

        # Calculate the sum row for the specified date range
        sum_row_custom = filtered_data[numeric_cols].sum(numeric_only=True)
        sum_row_custom['Date'] = 'Total'  # Label for the sum row
        
        # Append the sum row to the DataFrame
        filtered_data = filtered_data._append(sum_row_custom, ignore_index=True)

        # Drop the 'Month' column for the custom date range file if it exists
        if 'Month' in filtered_data.columns:
            filtered_data = filtered_data.drop(columns='Month')

        # Save the updated DataFrame with the sum row to Excel
        filtered_data.to_excel('custom_date_range_details.xlsx', index=False)
        print(f"Details for {start_date} to {end_date} saved to 'custom_date_range_details.xlsx'.")

def main():
    print("=== ANPR Fine Details Processing Tool ===")
    print("Please choose an option:")
    print("1. Generate Daily Details Report")
    print("2. Generate Monthly Summary Report")
    print("3. Generate Report for Custom Date Range")
    print("4. Exit")

    choice = input("Enter your choice (1-4): ")

    if choice == '4':
        print("Exiting the tool.")
        return

    directory = input("Enter the directory containing the CSV files: ")

    generate_daily = False
    generate_monthly = False
    start_date = None
    end_date = None

    if choice == '1':
        generate_daily = True
    elif choice == '2':
        generate_monthly = True
    elif choice == '3':
        start_date_str = input("Enter start date for custom report (YYYY-MM-DD): ")
        end_date_str = input("Enter end date for custom report (YYYY-MM-DD): ")
        try:
            start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
            end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
        except ValueError:
            print("Invalid date format. Please use YYYY-MM-DD.")
            return
    else:
        print("Invalid choice. Please select a valid option.")
        return

    # Call the processing function with the specified options
    process_all_csvs(directory, generate_daily, generate_monthly, start_date, end_date)

if __name__ == "__main__":
    main()
