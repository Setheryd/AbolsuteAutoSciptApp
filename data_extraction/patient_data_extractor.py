import win32com.client as win32
import pandas as pd
import os
import logging
from datetime import datetime
import sys

# Set up logging to both file and console
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# File handler
file_handler = logging.FileHandler('extract_eligible_patients.log')
file_handler.setLevel(logging.DEBUG)
file_formatter = logging.Formatter('%(asctime)s %(levelname)s:%(message)s')
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)

# Console handler
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.DEBUG)
console_formatter = logging.Formatter('%(asctime)s %(levelname)s:%(message)s')
console_handler.setFormatter(console_formatter)
logger.addHandler(console_handler)

class PatientDataExtractor:
    def __init__(self, base_path=None, password="abs$1018$B"):
        self.base_path = base_path or f"C:\\Users\\{os.getlogin()}\\OneDrive - Ability Home Health, LLC\\"
        self.password = password
        self.files_info = {
            "Absolute Patient Records.xlsm": "Absolute Operation",
            "Absolute Patient Records IHCC.xlsm": "IHCC",
            "Absolute Patient Records PERS.xlsm": "IHCC"
        }
        logging.debug(f"Initialized PatientDataExtractor with base_path: {self.base_path}")

    def find_file(self, base_path, filename, max_depth=5):
        """
        Optimized search for a file in a directory up to a specified depth using os.scandir for speed.
        """
        logging.debug(f"Searching for '{filename}' in '{base_path}' up to depth {max_depth}")
        
        def scan_directory(path, current_depth):
            if current_depth > max_depth:
                
                return None
            try:
                with os.scandir(path) as it:
                    for entry in it:
                        if entry.is_file() and entry.name.lower() == filename.lower():
                            logging.debug(f"Found file: {entry.path}")
                            return entry.path
                        elif entry.is_dir():
                            found_file = scan_directory(entry.path, current_depth + 1)
                            if found_file:
                                return found_file
            except PermissionError as e:
                logging.warning(f"PermissionError accessing {path}: {e}")
                return None
            except Exception as e:
                logging.error(f"Error scanning directory {path}: {e}")
                return None
        
        return scan_directory(base_path, 0)

    def extract_eligible_patients(self, output_to_csv=False, output_file=None):
        """
        Extract eligible patients from Excel files and optionally output the DataFrame as CSV.

        Args:
            output_to_csv (bool): If True, output the DataFrame as CSV.
            output_file (str): The file path to save the CSV. If None, print to stdout.
        Returns:
            pd.DataFrame or None: The extracted DataFrame, or None if no data was collected.
        """
        try:
            logging.debug(f"Starting extraction with base_path: {self.base_path}")
            if not os.path.exists(self.base_path):
                logging.error(f"Base path does not exist: {self.base_path}")
                return None

            collected_data = []
            excel = None
            workbooks = {}
            try:
                # Initialize Excel application
                logging.debug("Initializing Excel application...")
                excel = win32.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.ScreenUpdating = False
                excel.DisplayAlerts = False
                excel.EnableEvents = False
                excel.AskToUpdateLinks = False
                excel.AlertBeforeOverwriting = False
                logging.debug("Excel application started successfully.")

                # Open all workbooks
                for filename, required_subdir in self.files_info.items():
                    logging.debug(f"Processing file info: {filename} in subdir: {required_subdir}")
                    file_path = self.find_file(self.base_path, filename)
                    if file_path is None:
                        logging.warning(f"File not found: {filename}")
                        continue

                    # Verify the file is in the correct subdirectory
                    normalized_path = os.path.normpath(file_path)
                    if required_subdir.lower() not in normalized_path.lower():
                        logging.warning(f"File '{filename}' is not in the required subdirectory: '{required_subdir}'")
                        continue

                    logging.debug(f"Opening file: {file_path}")

                    try:
                        wb = excel.Workbooks.Open(
                            file_path,            # Filename
                            False,                # UpdateLinks
                            True,                 # ReadOnly
                            None,                 # Format
                            self.password,        # Password
                            '',                   # WriteResPassword
                            True,                 # IgnoreReadOnlyRecommended
                            None,                 # Origin
                            None,                 # Delimiter
                            False,                # Editable
                            False,                # Notify
                            None,                 # Converter
                            False,                # AddToMru
                            False,                # Local
                            0                     # CorruptLoad
                        )
                        workbooks[filename] = wb
                        logging.debug(f"Successfully opened workbook: {filename}")
                    except Exception as e:
                        logging.error(f"Error opening file '{file_path}': {e}")
                        continue

                # Process each workbook
                for filename, wb in workbooks.items():
                    logging.debug(f"Processing workbook: {filename}")
                    try:
                        ws = wb.Sheets("Patient Information")
                        used_range = ws.UsedRange
                        data = used_range.Value  # Read all data at once

                        if not data:
                            logging.warning(f"No data found in 'Patient Information' sheet in '{filename}'")
                            continue

                        # Assuming data is a tuple of tuples
                        data_rows = data[1:]  # Skip header row

                        for row in data_rows:
                            if filename == "Absolute Patient Records.xlsm":
                                # For "Absolute Patient Records.xlsm", use column J (10th column) for NOA start date
                                noa_start_date = row[9] if len(row) >= 10 else None  # Column J (10th)
                                discharge_date = row[11] if len(row) >= 12 else None  # Column L (12th)
                                if noa_start_date:
                                    patient_name = row[2] if len(row) >= 3 else ''  # Column C (3rd)
                                    collected_data.append({
                                        "Patient Name": patient_name,
                                        "First NOA Date": noa_start_date,
                                        "Discharge Date": discharge_date  # Replacing extra column with Discharge Date
                                    })
                            else:
                                # For other files, use column K (11th column) for NOA start date
                                noa_start_date = row[10] if len(row) >= 11 else None  # Column K (11th)
                                discharge_date = row[13] if len(row) >= 14 else None  # Column N (14th)
                                if noa_start_date:
                                    patient_name = row[3] if len(row) >= 4 else ''  # Column D (4th)
                                    collected_data.append({
                                        "Patient Name": patient_name,
                                        "First NOA Date": noa_start_date,
                                        "Discharge Date": discharge_date  # Renamed column
                                    })
                        logging.debug(f"Finished processing workbook: {filename}")
                    except Exception as e:
                        logging.error(f"Error processing workbook '{filename}': {e}")
                        continue

            except Exception as e:
                logging.error(f"An unexpected error occurred: {e}")
                return None
            finally:
                # Close all workbooks and quit Excel
                if workbooks:
                    logging.debug("Closing all open workbooks...")
                    for wb in workbooks.values():
                        try:
                            wb.Close(False)
                        except Exception as e:
                            logging.warning(f"Error closing workbook: {e}")
                if excel:
                    logging.debug("Quitting Excel application...")
                    excel.Quit()
                    del excel
                    logging.debug("Excel application closed.")

            if collected_data:
                # Create a DataFrame from the collected data
                df = pd.DataFrame(collected_data)
                logging.info("DataFrame created successfully.")

                # Convert date columns to strings in 'YYYY-MM-DD' format
                # This step ensures that the dates do not include time and timezone information
                def format_date(date_obj):
                    if isinstance(date_obj, datetime):
                        return date_obj.strftime('%Y-%m-%d')
                    elif isinstance(date_obj, str):
                        try:
                            # Attempt to parse the string to a datetime object
                            parsed_date = pd.to_datetime(date_obj, errors='coerce')
                            if pd.notnull(parsed_date):
                                return parsed_date.strftime('%Y-%m-%d')
                            else:
                                return date_obj  # Return as is if parsing fails
                        except Exception:
                            return date_obj  # Return as is if any exception occurs
                    else:
                        return date_obj  # Return as is for other types

                df['First NOA Date'] = df['First NOA Date'].apply(format_date)
                df['Discharge Date'] = df['Discharge Date'].apply(format_date)

                # Drop duplicates based on the relevant columns
                initial_count = len(df)
                df.drop_duplicates(inplace=True)
                final_count = len(df)
                logging.debug(f"Dropped {initial_count - final_count} duplicate rows.")

                if output_to_csv:
                    if output_file:
                        try:
                            df.to_csv(output_file, index=False)
                            logging.info(f"DataFrame saved to CSV file: {output_file}")
                        except Exception as e:
                            logging.error(f"Failed to save DataFrame to CSV: {e}")
                    else:
                        try:
                            df.to_csv(sys.stdout, index=False)
                            logging.info("DataFrame output to stdout.")
                        except Exception as e:
                            logging.error(f"Failed to output DataFrame to stdout: {e}")

                return df
            else:
                logging.info("No data collected.")
                return None

        except Exception as e:
            logging.error(f"Error in extract_eligible_patients: {e}")
            return None

if __name__ == "__main__":
    extractor = PatientDataExtractor()
    # Specify the output CSV file path if desired, e.g., "eligible_patients.csv"
    df = extractor.extract_eligible_patients(output_to_csv=True, output_file="eligible_patients.csv")
    if df is not None:
        logging.info(f"Extraction completed successfully. Total eligible patients: {len(df)}")
    else:
        logging.info("Extraction completed with no data collected.")
