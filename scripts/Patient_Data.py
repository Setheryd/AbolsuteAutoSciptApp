# extract_patients.py
import sys
from patient_data_extractor import PatientDataExtractor

def main():
    extractor = PatientDataExtractor()
    df = extractor.extract_eligible_patients(output_to_csv=True)  # Output to stdout by default
    if df is not None:
        # DataFrame is already printed to stdout by extract_eligible_patients
        pass
    else:
        sys.exit(1)  # Exit with error code if no data

if __name__ == "__main__":
    main()
