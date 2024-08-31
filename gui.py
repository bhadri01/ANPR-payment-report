import pandas as pd
from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkcalendar import DateEntry
from threading import Thread, Event

# Function to process the CSV files
def process_all_csvs(directory, generate_daily, generate_monthly, start_date, end_date, log_text, progress_var, stop_event):
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
    files = [f for f in os.listdir(directory) if f.endswith('.csv')]
    total_files = len(files)
    processed_files = 0

    for filename in files:
        if stop_event.is_set():
            log_text.insert(tk.END, "Processing canceled.\n")
            log_text.see(tk.END)
            root.update()
            return
        
        filepath = os.path.join(directory, filename)
        log_text.insert(tk.END, f"Processing file: {filename}\n")
        log_text.see(tk.END)
        root.update()  # Update the GUI dynamically

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
            log_text.insert(tk.END, f"Valid header not found in {filename}. Skipping this file.\n")
            log_text.see(tk.END)
            root.update()
            continue

        # Convert 'Payment Date' to datetime and normalize to just the date
        data['Payment Date'] = pd.to_datetime(data['Payment Date'], errors='coerce').dt.date

        # Handle cases where dates couldn't be converted (if any)
        if data['Payment Date'].isnull().any():
            log_text.insert(tk.END, f"Some dates in {filename} could not be converted and will be ignored.\n")
            log_text.see(tk.END)
            root.update()

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

        processed_files += 1
        progress_var.set((processed_files / total_files) * 100)
        log_text.insert(tk.END, f"Done processing file: {filename}\n")
        log_text.see(tk.END)
        root.update()

    # Create a 'Reports' directory inside the selected folder
    reports_directory = os.path.join(directory, "Reports")
    os.makedirs(reports_directory, exist_ok=True)

    # Calculate and append the sum row for the daily data
    if generate_daily:
        sum_row_daily = final_processed_data.sum(numeric_only=True)
        sum_row_daily['Date'] = 'Total'
        final_processed_data = final_processed_data._append(sum_row_daily, ignore_index=True)

        # Save the final aggregated daily DataFrame to Excel
        daily_report_path = os.path.join(reports_directory, 'ANPR_payment_details_daily.xlsx')
        final_processed_data.to_excel(daily_report_path, index=False)
        log_text.insert(tk.END, f"Daily details saved to '{daily_report_path}'.\n")
        log_text.see(tk.END)
        root.update()

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
        monthly_report_path = os.path.join(reports_directory, 'ANPR_payment_details_monthly.xlsx')
        month_wise_summary.to_excel(monthly_report_path, index=False)
        log_text.insert(tk.END, f"Monthly summary saved to '{monthly_report_path}'.\n")
        log_text.see(tk.END)
        root.update()

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
        custom_report_filename = f'ANPR_payment_details_{start_date}_to_{end_date}.xlsx'
        custom_report_path = os.path.join(reports_directory, custom_report_filename)
        filtered_data.to_excel(custom_report_path, index=False)
        log_text.insert(tk.END, f"Details for {start_date} to {end_date} saved to '{custom_report_path}'.\n")
        log_text.see(tk.END)
        root.update()

    # Show completion message
    messagebox.showinfo("Processing Complete", "All reports have been generated and saved in the 'Reports' folder.")

    # Enable the "Start Processing" button and disable the "Cancel" button after completion
    start_button.config(state=tk.NORMAL)
    cancel_button.config(state=tk.DISABLED)

def select_directory():
    directory = filedialog.askdirectory()
    directory_var.set(directory)

def toggle_date_fields():
    if custom_var.get():
        start_date_label.grid(row=5, column=0, sticky='w')
        start_date_entry.grid(row=5, column=1, padx=5)
        end_date_label.grid(row=6, column=0, sticky='w')
        end_date_entry.grid(row=6, column=1, padx=5)
    else:
        start_date_label.grid_remove()
        start_date_entry.grid_remove()
        end_date_label.grid_remove()
        end_date_entry.grid_remove()

def start_processing_thread():
    # Disable the "Start Processing" button and enable the "Cancel" button
    start_button.config(state=tk.DISABLED)
    cancel_button.config(state=tk.NORMAL)
    
    # Start the processing in a separate thread
    stop_event.clear()
    processing_thread = Thread(target=start_processing)
    processing_thread.start()

def start_processing():
    directory = directory_var.get()
    if not directory:
        messagebox.showerror("Error", "Please select a directory containing CSV files.")
        return

    generate_daily = daily_var.get()
    generate_monthly = monthly_var.get()
    start_date = None
    end_date = None

    if custom_var.get():
        try:
            start_date = datetime.strptime(start_date_entry.get(), '%Y-%m-%d').date()
            end_date = datetime.strptime(end_date_entry.get(), '%Y-%m-%d').date()
        except ValueError:
            messagebox.showerror("Error", "Invalid date format. Please use YYYY-MM-DD.")
            return

    process_all_csvs(directory, generate_daily, generate_monthly, start_date, end_date, log_text, progress_var, stop_event)

def cancel_processing():
    stop_event.set()

# GUI Setup
root = tk.Tk()
root.title("ANPR Fine Details Processing Tool")

# Directory selection
directory_var = tk.StringVar()
tk.Label(root, text="Select Directory:").grid(row=0, column=0, sticky='w')
tk.Entry(root, textvariable=directory_var, width=50).grid(row=0, column=1, padx=5)
tk.Button(root, text="Browse", command=select_directory).grid(row=0, column=2, padx=5)

# Option selections
daily_var = tk.BooleanVar()
monthly_var = tk.BooleanVar()
custom_var = tk.BooleanVar()

tk.Checkbutton(root, text="Generate Daily Details Report", variable=daily_var).grid(row=1, column=0, sticky='w')
tk.Checkbutton(root, text="Generate Monthly Summary Report", variable=monthly_var).grid(row=2, column=0, sticky='w')
tk.Checkbutton(root, text="Generate Report for Custom Date Range", variable=custom_var, command=toggle_date_fields).grid(row=3, column=0, sticky='w')

# Date range input
start_date_label = tk.Label(root, text="Start Date (YYYY-MM-DD):")
start_date_entry = DateEntry(root, date_pattern='yyyy-mm-dd')

end_date_label = tk.Label(root, text="End Date (YYYY-MM-DD):")
end_date_entry = DateEntry(root, date_pattern='yyyy-mm-dd')

toggle_date_fields()  # Initially hide date fields

# Log output
log_text = tk.Text(root, height=10, width=80)
log_text.grid(row=7, column=0, columnspan=3, pady=10)

# Progress bar
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.grid(row=8, column=0, columnspan=3, pady=10, sticky='we')

# Start and Cancel buttons
start_button = tk.Button(root, text="Start Processing", command=start_processing_thread)
start_button.grid(row=9, column=1, pady=10, sticky='ew')

cancel_button = tk.Button(root, text="Cancel", command=cancel_processing, state=tk.DISABLED)
cancel_button.grid(row=10, column=1, pady=10, sticky='ew')

# Event to control stopping the thread
stop_event = Event()

root.mainloop()
