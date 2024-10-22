import os
import glob
import pandas as pd

def find_target_directory(base_path, target_suffix):
    """
    Traverse the directory tree starting from base_path to find the directory
    that ends with the target_suffix.

    Parameters:
    - base_path (str): The base directory to start the search.
    - target_suffix (str): The directory suffix to match.

    Returns:
    - str: The full path to the target directory.

    Raises:
    - FileNotFoundError: If no matching directory is found.
    - Exception: If multiple matching directories are found.
    """
    target_suffix_parts = target_suffix.strip(os.sep).split(os.sep)
    suffix_length = len(target_suffix_parts)
    matching_directories = []

    for root, dirs, files in os.walk(base_path):
        # Split the current path into parts
        current_path_parts = os.path.normpath(root).split(os.sep)
        if len(current_path_parts) >= suffix_length:
            # Extract the last 'suffix_length' parts of the current path
            if current_path_parts[-suffix_length:] == target_suffix_parts:
                matching_directories.append(root)

    if not matching_directories:
        raise FileNotFoundError(f"No directory ending with '{target_suffix}' found under '{base_path}'.")

    if len(matching_directories) > 1:
        raise Exception(f"Multiple directories ending with '{target_suffix}' found under '{base_path}':\n" + "\n".join(matching_directories))

    return matching_directories[0]

def main():
    # Define the suffix path we're looking for
    target_suffix = os.path.join("Absolute Billing and Payroll", "Eligibility", "Eligibility Archive")

    # Dynamically construct the base path using the current user's login
    try:
        username = os.getlogin()
    except Exception as e:
        raise Exception("Unable to retrieve the current user's login name.") from e

    base_path = os.path.join("C:\\Users", username, "OneDrive - Ability Home Health, LLC")

    # Find the target directory
    target_directory = find_target_directory(base_path, target_suffix)

    # Get a list of all CSV files in the target directory
    csv_files = glob.glob(os.path.join(target_directory, "*.csv"))

    if not csv_files:
        raise FileNotFoundError(f"No CSV files found in directory: {target_directory}")

    # Find the newest CSV file based on modification time
    newest_csv = max(csv_files, key=os.path.getmtime)

    # Read the newest CSV into a DataFrame
    new_df = pd.read_csv(newest_csv)

    # Add the new DataFrame to an existing DataFrame
    try:
        # If existing_df already exists, append the new data
        existing_df = pd.concat([existing_df, new_df], ignore_index=True)
    except NameError:
        # If existing_df doesn't exist, initialize it with new_df
        existing_df = new_df.copy()

    # Optional: Save the updated DataFrame to a new CSV
    updated_csv_path = os.path.join(target_directory, "Updated_Patient_Count.csv")
    existing_df.to_csv(updated_csv_path, index=False)

    # Filter rows where Eligibility == 'Eligible'
    eligible_df = existing_df[existing_df['Eligibility'] == 'Eligible']

    # Count unique occurrences in 'Waiver/MCE' column
    waiver_mce_counts = eligible_df['Waiver/MCE'].value_counts().reset_index()
    waiver_mce_counts.columns = ['Waiver/MCE', 'Count']

    # Sort the DataFrame from largest to smallest count
    waiver_mce_counts = waiver_mce_counts.sort_values(by='Count', ascending=False).reset_index(drop=True)

    # Calculate the total count for percentage calculation
    total_count = waiver_mce_counts['Count'].sum()

    # Add a 'Percent of Total' column
    waiver_mce_counts['Percent of Total'] = (waiver_mce_counts['Count'] / total_count * 100).round(2)

    # Print the counts with percentages in comma-delimited (CSV) format
    print(waiver_mce_counts.to_csv(index=False))

if __name__ == "__main__":
    main()
