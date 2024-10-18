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
        self.base_path = base_path or f"C:\\Users\\{os.getlogin()}\\OneDrive - Ability Home Health, LLC\\Absolute Operation\\"
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

    def extract_caregivers(self, output_to_csv=False, output_file=None):
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
                        print(f"File found: {file_path}")  # Debugging print
                        logging.info(f"File found: {file_path}")

                    # Open the workbook and list sheets
                    try:
                        wb = excel.Workbooks.Open(file_path, False, True, None, self.password, '', True)
                        print([sheet.Name for sheet in wb.Sheets])  # Debugging print: check available sheets
                    except Exception as e:
                        logging.error(f"Error opening file {file_path}: {e}")
                        print(f"Error opening file {file_path}: {e}")  # Debugging print
                        continue

                    # Ensure sheet exists
                    try:
                        ws = wb.Sheets("Contractor_Employee")
                        used_range = ws.UsedRange
                        data = used_range.Value  # Read all data at once
                        # print(f"Used range data: {data}")  # Debugging print

                        if not data:
                            logging.warning(f"No data found in 'Contractor_Employee' sheet in '{filename}'")
                            print(f"No data found in 'Contractor_Employee' sheet in '{filename}'")  # Debugging print
                            continue
                    except Exception as e:
                        logging.error(f"Error reading sheet 'Contractor_Employee' in {filename}: {e}")
                        print(f"Error reading sheet 'Contractor_Employee' in {filename}: {e}")  # Debugging print
                        continue

                    # Extract data from columns C, H, J
                    for row in data[2:]:  # Skip header row
                        # print(f"Row data: {row}")  # Debugging print
                        contractor_name = row[2] if len(row) >= 3 else None  # Column C (3rd)
                        date_of_hire = row[7] if len(row) >= 8 else None  # Column H (8th)
                        term_date = row[9] if len(row) >= 10 else None  # Column J (10th)

                        # Debugging prints
                        # print(f"Contractor Name: {contractor_name}, Date of Hire: {date_of_hire}, Term Date: {term_date}")

                        if contractor_name:
                            collected_data.append({
                                "Contractor Name (C)": contractor_name,
                                "Date of Hire (H)": date_of_hire,
                                "Term Date (J)": term_date
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

                df["Date of Hire (H)"] = df["Date of Hire (H)"].apply(format_date)
                df["Term Date (J)"] = df["Term Date (J)"].apply(format_date)
                
                if output_to_csv:
                    if output_file:
                        df.to_csv(output_file, index=False)
                        logging.info(f"DataFrame saved to CSV file: {output_file}")
                    else:
                        df.to_csv(sys.stdout, index=False)
                        logging.info("DataFrame output to stdout.")

                return df
            else:
                logging.info("No data collected.")
                print("No data collected.")  # Debugging print
                return None

        except Exception as e:
            logging.error(f"Error in caregiver_data_extractor: {e}")
            print(f"Error: {e}")  # Debugging print
            return None
        
if __name__ == "__main__":
    extractor = CaregiverDataExtractor()
    df = extractor.extract_caregivers()
    if df is not None:
        print(df)
    else:
        print("No eligible patient data was extracted.")
