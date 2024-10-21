import sys
import pandas as pd
from employee_records_data_extractor import EmployeeRecordsExtractor

def main():
    extractor = EmployeeRecordsExtractor()
    
    # Extract the DataFrame without using output_to_csv
    try:
        df = extractor.process_scheduling_files()  # Removed the output_to_csv argument
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
        # Step 1: Filter rows where the "Caregiver" column is not "Assigned Hrs."
        filtered_df = df[df['Caregiver'] != 'Assigned Hrs.']

        # Step 2: Identify columns that represent dates and remove duplicates
        date_columns = []
        for col in filtered_df.columns:
            try:
                # Try to convert column name to datetime; if successful, it's a date column
                pd.to_datetime(col, format='%Y-%m-%d', errors='raise')
                if col not in date_columns:
                    date_columns.append(col)  # Only keep unique date columns
            except (ValueError, TypeError):
                # If conversion fails, skip this column
                continue

        # Step 3: Sum the hours for each unique date column (excluding "Assigned Hrs.")
        sum_by_month = filtered_df[date_columns].sum()

        # Step 4: Group the columns by their month
        sum_by_month.index = pd.to_datetime(sum_by_month.index)
        grouped_sum_by_month = sum_by_month.groupby(sum_by_month.index.to_period("M")).sum()

        # Step 5: Calculate 'Assigned Hours' where Caregiver is "Assigned Hrs."
        assigned_hours_df = df[df['Caregiver'] == 'Assigned Hrs.']
        assigned_hours_sum = assigned_hours_df[date_columns].sum()

        # Group assigned hours by the same monthly periods
        assigned_hours_sum.index = pd.to_datetime(assigned_hours_sum.index)
        grouped_assigned_hours_sum = assigned_hours_sum.groupby(assigned_hours_sum.index.to_period("M")).sum()

        # Step 6: Convert the Series to DataFrame and rename the columns
        grouped_sum_by_month_df = grouped_sum_by_month.reset_index()
        grouped_sum_by_month_df.columns = ['Month', 'Completed Hours']

        # Adding the 'Assigned Hours' column to the final DataFrame
        grouped_sum_by_month_df['Assigned Hours'] = grouped_assigned_hours_sum.values

        # Step 7: Add a new column for Completed Hours / Assigned Hours
        grouped_sum_by_month_df['Utilization Rate'] = grouped_sum_by_month_df['Completed Hours'] / grouped_sum_by_month_df['Assigned Hours']

        # Handle division by zero or NaN values gracefully
        grouped_sum_by_month_df['Utilization Rate'] = grouped_sum_by_month_df['Utilization Rate'].replace([float('inf'), -float('inf')], 0).fillna(0)

        # Step 8: Format the 'Hours Ratio' as a percentage with two decimal places
        grouped_sum_by_month_df['Utilization Rate'] = (grouped_sum_by_month_df['Utilization Rate'] * 100).round(2).astype(str) + '%'

        # Step 9: Output the CSV string (which is delimited data)
        csv_output = grouped_sum_by_month_df.to_csv(index=False)
        print(csv_output)

    else:
        sys.exit(1)  # Exit with error code if no data

if __name__ == "__main__":
    main()
