import sys
import pandas as pd
from patient_data_extractor import PatientDataExtractor

def main():
    extractor = PatientDataExtractor()
    
    # Extract the DataFrame directly
    df = extractor.extract_eligible_patients()  # Removed output_to_csv argument

    if df is not None:
        # Convert 'First NOA Date' and 'Discharge Date' to datetime for proper handling
        df['First NOA Date'] = pd.to_datetime(df['First NOA Date'], errors='coerce')
        df['Discharge Date'] = pd.to_datetime(df['Discharge Date'], errors='coerce')

        # Determine the date range (from the earliest month to the current month)
        start_date = df['First NOA Date'].min().replace(day=1)
        end_date = pd.Timestamp.today().replace(day=1)

        # Create a list of months between the start and end date
        all_months = pd.date_range(start=start_date, end=end_date, freq='MS')

        # Create an empty list to store the count of active patients per month
        active_patient_counts = []

        # Loop through each month and count the number of active patients
        for month in all_months:
            active_patients = df[
                (df['First NOA Date'] <= month) &
                ((df['Discharge Date'].isna()) | (df['Discharge Date'] >= month))
            ]
            active_patient_counts.append({'Month-Year': month.strftime('%B-%Y'), 'Active Patients': len(active_patients)})

        # Convert the list of results into a DataFrame
        df_active_patients_by_month = pd.DataFrame(active_patient_counts)

        # Output the DataFrame to stdout
        print(df_active_patients_by_month.to_csv(index=False))

    else:
        sys.exit(1)  # Exit with error code if no data

if __name__ == "__main__":
    main()
