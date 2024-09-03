import os
from datetime import datetime
import pandas as pd

def get_recent_slddrw_files(root_dir, start_date, end_date):
    # Initialize a list to store file paths and their modification times
    files_with_times = []

    # Walk through all directories and subdirectories
    for dirpath, dirnames, filenames in os.walk(root_dir):
        # Convert the folder name to lowercase for case-insensitive comparison
        folder_name_lower = os.path.basename(dirpath).lower()

        # Skip directories named "old rev" or "down rev"
        if folder_name_lower in ["old rev", "down rev", "old revs", "old revs AAP", "old revisions"]:
            continue
        
        for filename in filenames:
            # Check if the file has a .SLDDRW extension
            if filename.lower().endswith('.slddrw'):
                file_path = os.path.join(dirpath, filename)
                
                try:
                    # Get the file's modification time and creation time
                    file_mtime = os.path.getmtime(file_path)
                    file_ctime = os.path.getctime(file_path)
                    file_time = max(file_mtime, file_ctime)
                    file_time_datetime = datetime.fromtimestamp(file_time)

                    # Check if the file's modification time is within the specified date range
                    if start_date <= file_time_datetime <= end_date:
                        # Strip off the root directory part from the file path
                        relative_path = os.path.relpath(file_path, root_dir)
                        # Append the relative path and its modification/creation time to the list
                        files_with_times.append((relative_path, file_time_datetime))
                except OSError:
                    # Handle the case where the file might be inaccessible
                    print(f"Could not access {file_path}. Skipping.")

    # Sort the files by their modification/creation time in descending order (most recent first)
    files_with_times.sort(reverse=True, key=lambda x: x[1])

    # Convert the list to a DataFrame
    df = pd.DataFrame(files_with_times, columns=["File Path", "Date Modified"])

    # Export the DataFrame to an Excel file
    output_file = 'slddrw_files_filtered.xlsx'
    df.to_excel(output_file, index=False)
    
    print(f"Exported to {output_file}")

# Get user input for date range
start_date_str = input("Enter the start date (YYYY-MM-DD): ")
end_date_str = input("Enter the end date (YYYY-MM-DD): ")

# Convert user input to datetime objects
start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
end_date = datetime.strptime(end_date_str, "%Y-%m-%d")

# Specify the directory you want to search
root_directory = r'*Your  directory here'

# Run the function with the user-provided date range
get_recent_slddrw_files(root_directory, start_date, end_date)