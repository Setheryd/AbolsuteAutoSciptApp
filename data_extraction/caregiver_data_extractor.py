import win32com.client as win32
import pandas as pd
import os
import logging
from datetime import datetime
import sys

# Set up logging to a file
logging.basicConfig(
    filename='extract_eligible_patients.log',
    level=logging.DEBUG,
    format='%(asctime)s %(levelname)s:%(message)s'
)

class CaregiverDataExtractor:
    def __init__(self, base_path=None, password="abs$1004$N"):
        self.base_path = base_path or f"C:\\Users\\{os.getlogin()}\\OneDrive - Ability Home Health, LLC\\"
        self.password = password
        self.files_info = {
            "Absolute Employee Demographics.xlsm": "Employee Demographics File",
        }

    def find_file(self, base_path, filename, max_depth=5):
        """
        Optimized search for a file in a directory up to a specified depth using os.scandir for speed.
        """
        logging.debug(f"Searching for {filename} in {base_path} up to depth {max_depth}")
        
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
                logging.warning(f"PermissionError: {e}")
                return None
        
        return scan_directory(base_path, 0)

    def extract_caregivers(self):
        try:
            logging.debug(f"Base path: {self.base_path}")
            if not os.path.exists(self.base_path):
                logging.error(f"Base path does not exist: {self.base_path}")
                print(f"Base path does not exist: {self.base_path}")  # Debugging print
                return None

            collected_data = []
            try:
                # Use DispatchEx to create a new Excel instance
                excel = win32.DispatchEx("Excel.Application")
                excel.Visible = False
                logging.debug("Excel application started successfully.")
            except Exception as e:
                logging.error(f"Failed to create Excel application: {e}")
                print(f"Failed to create Excel application: {e}")  # Debugging print
                return None

            workbooks = {}
            try:
                # Open all workbooks
                for filename, required_subdir in self.files_info.items():
                    file_path = self.find_file(self.base_path, filename)
                    if file_path is None:
                        logging.warning(f"File not found: {filename}")
                        print(f"File not found: {filename}")  # Debugging print
                        continue
                    else:
                        logging.info(f"File found: {file_path}")

                    # Open the workbook and list sheets
                    try:
                        wb = excel.Workbooks.Open(file_path, False, True, None, self.password, '', True)
                        workbooks[filename] = wb  # Keep track of opened workbooks
                    except Exception as e:
                        logging.error(f"Error opening file {file_path}: {e}")
                        print(f"Error opening file {file_path}: {e}")  # Debugging print
                        continue

                    # Ensure sheet exists
                    try:
                        ws = wb.Sheets("Contractor_Employee")
                        used_range = ws.UsedRange
                        data = used_range.Value  # Read all data at once

                        if not data:
                            logging.warning(f"No data found in 'Contractor_Employee' sheet in '{filename}'")
                            print(f"No data found in 'Contractor_Employee' sheet in '{filename}'")  # Debugging print
                            continue
                    except Exception as e:
                        logging.error(f"Error reading sheet 'Contractor_Employee' in {filename}: {e}")
                        print(f"Error reading sheet 'Contractor_Employee' in {filename}: {e}")  # Debugging print
                        continue

                    # Ensure there are at least two rows (assuming row 2 has headers)
                    if len(data) < 2:
                        logging.warning(f"Not enough rows in 'Contractor_Employee' sheet in '{filename}'")
                        print(f"Not enough rows in 'Contractor_Employee' sheet in '{filename}'")  # Debugging print
                        continue

                    # Extract header row (assuming row 2 is at index 1)
                    header_row = data[1]  # Row 2
                    desired_headers = ["Last, First M", "DateofHire", "Termination date"]
                    header_to_index = {}

                    for idx, header in enumerate(header_row):
                        if header in desired_headers:
                            header_to_index[header] = idx

                    # Check if all desired headers are found
                    missing_headers = [h for h in desired_headers if h not in header_to_index]
                    if missing_headers:
                        logging.error(f"Missing headers {missing_headers} in 'Contractor_Employee' sheet in '{filename}'")
                        print(f"Missing headers {missing_headers} in 'Contractor_Employee' sheet in '{filename}'")  # Debugging print
                        continue

                    logging.debug(f"Header indices: {header_to_index}")

                    # Extract data from rows starting after the header
                    for row_num, row in enumerate(data[2:], start=3):  # Starting at row 3
                        # Handle cases where the row might be shorter than expected
                        try:
                            contractor_name = row[header_to_index["Last, First M"]] if len(row) > header_to_index["Last, First M"] else None
                            date_of_hire = row[header_to_index["DateofHire"]] if len(row) > header_to_index["DateofHire"] else None
                            term_date = row[header_to_index["Termination date"]] if len(row) > header_to_index["Termination date"] else None
                        except Exception as e:
                            logging.warning(f"Error accessing data in row {row_num} of '{filename}': {e}")
                            contractor_name, date_of_hire, term_date = None, None, None

                        if contractor_name:
                            collected_data.append({
                                "Contractor Name": contractor_name,
                                "Date of Hire": date_of_hire,
                                "Term Date": term_date
                            })

            finally:
                # Close all workbooks and Excel application
                for wb in workbooks.values():
                    wb.Close(False)
                excel.Quit()
                del excel
                logging.debug("Excel application closed.")

            if collected_data:
                # Create a DataFrame from the collected data
                df = pd.DataFrame(collected_data)
                logging.info("DataFrame created successfully.")
                
                # Convert date columns to strings in 'YYYY-MM-DD' format
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

                df["Date of Hire"] = df["Date of Hire"].apply(format_date)
                df["Term Date"] = df["Term Date"].apply(format_date)
                
                return df
            else:
                logging.info("No data collected.")
                print("No data collected.")  # Debugging print
                return None

        except Exception as e:
            logging.error(f"Error in caregiver_data_extractor: {e}")
            print(f"Error: {e}")  # Debugging print
            return None
        
def main():
    extractor = CaregiverDataExtractor()
    df = extractor.extract_caregivers()
    if df is not None:
        # Output DataFrame in CSV format without index
        print(df.to_csv(index=False))
    else:
        print("No eligible patient data was extracted.")

if __name__ == "__main__":
    main()
