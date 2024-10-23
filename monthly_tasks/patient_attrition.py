# patient_attrition.py

import pandas as pd
import os
import sys
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import MaxNLocator


def get_resource_path(relative_path):
    """Get the absolute path to the resource, works for PyInstaller executable."""
    try:
        # PyInstaller creates a temporary folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # If not running as an executable, use the current script directory
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)


# Get the parent directory using get_resource_path
parent_dir = get_resource_path(os.path.join(os.pardir))

# Add the data_extraction directory to the system path
sys.path.append(os.path.join(parent_dir, "data_extraction"))

from patient_data_extractor import PatientDataExtractor


class ChurnAttritionAnalyzer:
    def __init__(self, extractor: PatientDataExtractor):
        self.extractor = extractor
        # Removed report_filename and outliers_filename since we are not saving files

    def load_data(self):
        """
        Load the patient data using the extractor.

        Returns:
            pd.DataFrame or None
        """
        df = self.extractor.extract_eligible_patients()
        if df is not None:
            # Convert date columns from strings to datetime objects
            df["First NOA Date"] = pd.to_datetime(df["First NOA Date"], errors="coerce")
            df["Discharge Date"] = pd.to_datetime(df["Discharge Date"], errors="coerce")
            # Drop records with invalid First NOA Date (start date)
            df = df.dropna(subset=["First NOA Date"])
            # Ensure no records have empty or null 'Patient Name'
            df = df.dropna(subset=["Patient Name"])
            return df
        else:
            print("No eligible patient data was extracted.")
            return None

    def generate_monthly_report(self, df, report_month=None):
        """
        Generate a churn and attrition rate report for a specific month.

        Args:
            df (pd.DataFrame): DataFrame containing patient data.
            report_month (datetime, optional): The month for which to generate the report. Defaults to last month.

        Returns:
            dict: A dictionary containing the report metrics.
        """
        if report_month is None:
            today = datetime.today()
            report_month = pd.to_datetime(
                datetime(today.year, today.month, 1)
            ) - pd.DateOffset(months=1)
        else:
            report_month = pd.to_datetime(report_month)

        # Define the start and end of the report month
        start_of_month = report_month.replace(day=1)
        end_of_month = (start_of_month + pd.DateOffset(months=1)) - pd.DateOffset(
            days=1
        )

        # Ensure start_of_month and end_of_month are offset-naive
        if start_of_month.tzinfo is not None:
            start_of_month = start_of_month.tz_localize(None)
        if end_of_month.tzinfo is not None:
            end_of_month = end_of_month.tz_localize(None)

        # Filter active patients at the start and end of the month
        active_start = df[
            (df["First NOA Date"] <= start_of_month)
            & ((df["Discharge Date"].isna()) | (df["Discharge Date"] > start_of_month))
        ]
        active_end = df[
            (df["First NOA Date"] <= end_of_month)
            & ((df["Discharge Date"].isna()) | (df["Discharge Date"] > end_of_month))
        ]

        # Calculate the number of active patients
        active_start_count = active_start["Patient Name"].nunique()
        active_end_count = active_end["Patient Name"].nunique()

        # Calculate net change
        net_change = active_end_count - active_start_count

        # Calculate active patients per month for average net change
        months = pd.period_range(
            df["First NOA Date"].min().to_period("M"),
            end_of_month.to_period("M"),
            freq="M",
        )
        active_counts = []
        for month in months:
            month_start = month.to_timestamp()
            month_end = (month_start + pd.DateOffset(months=1)) - pd.DateOffset(days=1)
            active = df[
                (df["First NOA Date"] <= month_end)
                & ((df["Discharge Date"].isna()) | (df["Discharge Date"] > month_end))
            ]
            active_counts.append(active["Patient Name"].nunique())

        # Calculate monthly net changes
        net_changes = [
            active_counts[i] - active_counts[i - 1]
            for i in range(1, len(active_counts))
        ]
        average_net_change = pd.Series(net_changes).mean() if net_changes else 0

        # Calculate discharges in the report month
        discharges = df[
            (df["Discharge Date"] >= start_of_month)
            & (df["Discharge Date"] <= end_of_month)
        ]
        num_discharges = discharges["Patient Name"].nunique()

        # Calculate new patients in the report month
        new_patients = df[
            (df["First NOA Date"] >= start_of_month)
            & (df["First NOA Date"] <= end_of_month)
        ]["Patient Name"].nunique()

        # Calculate churn and attrition rates based on new definitions
        churn = (
            (num_discharges / active_end_count) * 100 if active_end_count else 0
        )
        # Attrition rate: number of discharges during the month divided by average number of patients
        average_patients = (active_start_count + active_end_count) / 2
        attrition = (num_discharges / average_patients) * 100 if average_patients else 0

        # Prepare report
        report = {
            "Report Month": report_month.strftime("%B %Y"),
            "Starting Patient Count": active_start_count,
            "Ending Patient Count": active_end_count,
            "New Patients": new_patients,  # Added new patients
            "Discharged Patients": num_discharges,  # Added discharged patients
            "Net Change": net_change,
            "Average Net Change": round(average_net_change, 2),
            "Churn Rate (%)": round(churn, 2),
            "Attrition Rate (%)": round(attrition, 2),
        }

        return report

    def generate_all_monthly_reports(self, df):
        """
        Generate monthly reports for all months in the data.

        Args:
            df (pd.DataFrame): DataFrame containing patient data.

        Returns:
            pd.DataFrame: DataFrame containing all monthly reports.
        """
        months = pd.period_range(
            df["First NOA Date"].min().to_period("M"),
            pd.Timestamp.today().to_period("M"),
            freq="M",
        )
        report_data = []
        for month in months:
            report = self.generate_monthly_report(df, report_month=month.to_timestamp())
            report_data.append(report)
        report_df = pd.DataFrame(report_data)
        return report_df

    def get_csv_string(self, report_df):
        """
        Convert the report DataFrame to a comma-delimited string without indexing.

        Args:
            report_df (pd.DataFrame): DataFrame containing monthly reports.

        Returns:
            str: Comma-delimited string of the report data.
        """
        return report_df.to_csv(index=False)

    def generate_charts(self, report_df, output_dir="../charts"):
        """
        Generate charts showing churn and attrition rates over time and save as an image file.

        Args:
            report_df (pd.DataFrame): DataFrame containing monthly reports.
            output_dir (str): Directory to save the chart image.

        Returns:
            str: Absolute path to the saved chart image.
        """
        # Ensure the output directory exists
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Convert 'Report Month' to datetime for plotting
        report_df["Report Month Date"] = pd.to_datetime(
            report_df["Report Month"], format="%B %Y"
        )

        # Plot Churn and Attrition Rates
        plt.figure(figsize=(14, 7))
        plt.plot(
            report_df["Report Month Date"],
            report_df["Churn Rate (%)"],
            marker="o",
            label="Churn Rate (%)",
            color="#006400",  # Dark green for Churn Rate
            linewidth=2
        )
        plt.plot(
            report_df["Report Month Date"],
            report_df["Attrition Rate (%)"],
            marker="o",
            label="Attrition Rate (%)",
            color="#32CD32",  # Lighter green for Attrition Rate
            linewidth=2
        )

        plt.xlabel("Month")
        plt.ylabel("Rate (%)")
        plt.title("Churn and Attrition Rates Over Time")
        plt.legend()

        # Set major ticks to every 2 months
        ax = plt.gca()
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))  # Every 2 months
        ax.xaxis.set_major_formatter(
            mdates.DateFormatter("%B %Y")
        )  # e.g., January 2023

        # Rotate date labels for better readability
        plt.xticks(rotation=45)

        plt.tight_layout()

        # Define absolute path for the chart image
        chart_filename = get_resource_path(
            os.path.join(output_dir, "churn_attrition_chart.png")
        )

        # Save the figure
        plt.savefig(chart_filename)
        plt.close()  # Close the figure to free memory

        return chart_filename

    def run_analysis(self):
        """
        Run the complete churn and attrition analysis.
        """
        df = self.load_data()
        if df is None:
            print("No data available for analysis.")
            return

        # Generate all monthly reports
        report_df = self.generate_all_monthly_reports(df)

        # Instead of saving to CSV, get the CSV string
        csv_string = self.get_csv_string(report_df)

        print(csv_string)

        # Generate charts and get the chart filename
        chart_filename = self.generate_charts(report_df)

        # Optionally, you can return the report and chart filename
        return report_df, chart_filename


def main():
    extractor = PatientDataExtractor()
    analyzer = ChurnAttritionAnalyzer(extractor)
    analyzer.run_analysis()


if __name__ == "__main__":
    main()
