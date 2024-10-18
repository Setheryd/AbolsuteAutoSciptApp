import sys
import pandas as pd
from datetime import datetime
from caregiver_data_extractor import CaregiverDataExtractor

def main():
    extractor = CaregiverDataExtractor()
    
    # Extract the DataFrame without using output_to_csv
    try:
        df = extractor.extract_caregivers()  # Removed the output_to_csv argument
    except pd.errors.ParserError as e:
        print(f"Error parsing data: {e}")
        sys.exit(1)
    except pd.errors.EmptyDataError as e:
        print(f"No data found in the file: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

    if df is not None:
        # Ensure DataFrame is not empty
        if df.empty:
            print("The DataFrame is empty.")
            sys.exit(1)  # Exit with error code if no data

        # Convert 'Date of Hire' and 'Term Date' to datetime for proper handling
        df["Date of Hire (H)"] = pd.to_datetime(df["Date of Hire (H)"], errors='coerce')
        df["Term Date (J)"] = pd.to_datetime(df["Term Date (J)"], errors='coerce')

        # Determine the date range (from the earliest month to the current month)
        start_date = df["Date of Hire (H)"].min().replace(day=1)
        end_date = pd.Timestamp.today().replace(day=1)

        # Create a list of months between the start and end date
        all_months = pd.date_range(start=start_date, end=end_date, freq='MS')

        # Create an empty list to store the count of active caregivers per month
        active_caregiver_counts = []

        # Loop through each month and count the number of active caregivers
        for month in all_months:
            active_caregivers = df[
                (df["Date of Hire (H)"] <= month) & 
                ((df["Term Date (J)"].isna()) | (df["Term Date (J)"] >= month))
            ]
            active_caregiver_counts.append({'Month-Year': month.strftime('%B-%Y'), 'Active Caregivers': len(active_caregivers)})

        # Convert the list of results into a DataFrame
        df_active_caregivers_by_month = pd.DataFrame(active_caregiver_counts)

        # Output the DataFrame in CSV format (as a string)
        csv_output = df_active_caregivers_by_month.to_csv(index=False)

        # Output the CSV string (which is delimited data)
        print(csv_output)

    else:
        sys.exit(1)  # Exit with error code if no data

if __name__ == "__main__":
    main()
