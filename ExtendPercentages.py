import os
import glob
import re
import pandas as pd

# Create an empty DataFrame to store the results
result_df = pd.DataFrame(columns=["Year", "Month", "Number of Rows", "Count of Extends", "Percentage of Extends"])

# Define the year folders
year_folders = ["2021", "2022", "2023"]

# Loop through the year folders
for year_folder in year_folders:
    # Create the full path to the year folder
    year_folder_path = os.path.join("path/to/year/folder", year_folder)

    # Get the list of subfolders within the year folder
    subfolders = next(os.walk(year_folder_path))[1]

    # Loop through the subfolders
    for subfolder in subfolders:
        # Extract the month number from the subfolder name using regular expressions
        month_match = re.search(r"(\d+)", subfolder)
        if month_match:
            month = int(month_match.group(1))
        else:
            # If no month number is found, skip the folder
            continue

        # Create the full path to the subfolder
        subfolder_path = os.path.join(year_folder_path, subfolder)

        # Create the file path pattern to search for files
        file_pattern = os.path.join(subfolder_path, "*.xlsx")  # Adjust the pattern as per your file naming convention

        # Get the list of file paths that match the pattern
        file_paths = glob.glob(file_pattern)

        # Sort the file paths by modification time (latest file first)
        file_paths = sorted(file_paths, key=os.path.getmtime, reverse=True)

        # Select the latest file path
        latest_file_path = file_paths[0] if file_paths else None

        # Proceed if a file was found
        if latest_file_path:
            # Read the Excel file into a dictionary of DataFrames (one DataFrame per sheet)
            xls = pd.read_excel(latest_file_path, sheet_name=None)

            # Initialize variables to track the selected sheet and count of "extend" values
            selected_sheet = None
            max_extend_count = 0

            # Iterate over the sheets in the Excel file
            for sheet_name, df in xls.items():
                # Skip the current sheet if "Extend" column is not found
                if "Extend" not in df.columns:
                    continue

                # Convert "Extend" column values to strings and lowercase for case-insensitive matching
                df["Extend"] = df["Extend"].astype(str).str.lower()

                # Calculate the count of "extend" values
                count_extends = df["Extend"].eq("extend").sum()

                # Check if the count of "extend" values is the highest so far
                if count_extends > max_extend_count:
                    max_extend_count = count_extends
                    selected_sheet = sheet_name

            # Check if a sheet with "Extend" column was found
            if selected_sheet:
                # Read the selected sheet into a DataFrame
                df = pd.read_excel(latest_file_path, sheet_name=selected_sheet)

                # Convert "Extend" column values to strings and lowercase for case-insensitive matching
                df["Extend"] = df["Extend"].astype(str).str.lower()

                # Calculate the count of "extend" values
                count_extends = df["Extend"].eq("extend").sum()

                # Get the number of rows in the DataFrame
                num_rows = df.shape[0]

                # Append the results to the result DataFrame
                result_df = result_df.append(
                    {"Year": int(year_folder), "Month": f"{year_folder} - {month}", "Number of Rows": num_rows,
                     "Count of Extends": count_extends, "Percentage of Extends": count_extends / num_rows * 100},
                    ignore_index=True)

# Print the final result
print(result_df)

# Save the result_df DataFrame as a CSV file
result_df.to_csv('path/to/save/results.csv', index=False)
