import logging
from billing_files_extractor import BillingFilesDataExtractor  # Adjust the import path as needed
import pandas as pd

def setup_logging():
    """
    Configures logging to output to both console and a log file.
    """
    logging.basicConfig(
        level=logging.DEBUG,  # Set to DEBUG to capture detailed logs
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("billing_extractor.log"),  # Log to file
            logging.StreamHandler()  # Log to console
        ]
    )

def count_abs_in_medical_record(df):
    """
    Counts the number of rows containing 'abs' in the 'Medical Record Number' column
    for each source file.

    Parameters:
        df (pd.DataFrame): The DataFrame containing billing data.

    Returns:
        pd.DataFrame: A DataFrame with 'Source File' and corresponding 'abs_Count'.
    """
    # Check if necessary columns exist
    required_columns = ['Medical Record Number', 'Source File']
    for col in required_columns:
        if col not in df.columns:
            logging.error(f"Required column '{col}' not found in DataFrame.")
            raise KeyError(f"Required column '{col}' not found in DataFrame.")

    # Ensure 'Medical Record Number' is of string type
    df['Medical Record Number'] = df['Medical Record Number'].astype(str)

    # Create a boolean mask where 'Medical Record Number' contains 'abs' (case-insensitive)
    mask = df['Medical Record Number'].str.contains('abs', case=False, na=False)

    # Apply the mask to filter relevant rows
    filtered_df = df[mask]

    # Group by 'Source File' and count the occurrences
    counts = filtered_df.groupby('Source File').size().reset_index(name='abs_Count')

    return counts

def integrate_counts(df, counts_df):
    """
    Integrates the counts DataFrame into the main DataFrame by adding a new column
    that maps the 'abs_Count' to each row based on 'Source File'.

    Parameters:
        df (pd.DataFrame): The main DataFrame containing billing data.
        counts_df (pd.DataFrame): The DataFrame containing counts per 'Source File'.

    Returns:
        pd.DataFrame: The integrated DataFrame with an additional 'abs_Count' column.
    """
    # Merge the counts back into the main DataFrame
    integrated_df = pd.merge(df, counts_df, on='Source File', how='left')

    # Replace NaN counts with 0 (if there were no 'abs' entries for a Source File)
    integrated_df['abs_Count'] = integrated_df['abs_Count'].fillna(0).astype(int)

    return integrated_df

def main():
    """
    Main function to execute the extraction and counting process.
    """
    setup_logging()
    logging.info("Starting the billing files extraction and counting process.")

    try:
        # Instantiate the extractor
        extractor = BillingFilesDataExtractor()
        logging.info("Initialized BillingFilesDataExtractor.")

        # Process billing files to get the DataFrame
        df = extractor.process_billing_files()

        if df is not None and not df.empty:
            logging.info("Successfully extracted billing data.")

            # Perform the count
            counts_df = count_abs_in_medical_record(df)
            logging.info("Counting of 'abs' in 'Medical Record Number' completed.")

            if not counts_df.empty:
                # Integrate the counts into the main DataFrame
                integrated_df = integrate_counts(df, counts_df)
                logging.info("Integrated 'abs_Count' into the main DataFrame.")

                # Display the integrated DataFrame
                print("Integrated DataFrame with 'abs_Count':")
                print(integrated_df.head())  # Displaying only the first few rows for brevity

                # If you want to see the counts separately, you can still print counts_df
                print(counts_df.to_string(index=False))

                # Optionally, return the integrated DataFrame for further processing
                # return integrated_df

            else:
                # If there are no 'abs' records, add a column with 0 counts
                df['abs_Count'] = 0
                print("No records containing 'abs' were found in the 'Medical Record Number' column.")
                logging.info("No 'abs' records found in the 'Medical Record Number' column.")
                print("Integrated DataFrame with 'abs_Count':")
                print(df.head())  # Displaying only the first few rows for brevity

        else:
            print("No eligible patient data was extracted.")
            logging.info("No eligible patient data was extracted from billing files.")

    except FileNotFoundError as fnf_error:
        logging.error(f"File not found error: {fnf_error}")
        print(f"File not found error: {fnf_error}")
    except KeyError as key_error:
        logging.error(f"Key error: {key_error}")
        print(f"Key error: {key_error}")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
