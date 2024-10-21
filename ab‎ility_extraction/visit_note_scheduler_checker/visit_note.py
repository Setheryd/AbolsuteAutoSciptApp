from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import os
import time
from datetime import datetime, timedelta
from pathlib import Path
import shutil

class ReportDownloader:
    def __init__(self, target_folder = None):
        # Get the system's default download folder
        self.download_folder = self.get_default_download_folder()
        self.target_folder = target_folder or f"C:\\Users\\{os.getlogin()}\\Downloads"  # Set your target folder for moved files
        self.dfs = []  # To store the dataframes after loading

        # Set Chrome options to use the default download folder
        chrome_options = Options()
        prefs = {"download.default_directory": self.download_folder}
        chrome_options.add_experimental_option("prefs", prefs)

        # Initialize Chrome WebDriver
        self.driver = webdriver.Chrome(options=chrome_options)

    # Function to get previous month range
    def get_previous_month_range(self):
        today = datetime.now()
        first_day_of_current_month = today.replace(day=1)
        last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
        first_day_of_previous_month = last_day_of_previous_month.replace(day=1)
        return first_day_of_previous_month.strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d')

    # Function to click buttons with fallback
    def click_button_with_fallback(self):
        try:
            button = self.driver.find_element(By.ID, "csvDataExport")
            button.click()
        except Exception:
            try:
                button = self.driver.find_element(By.ID, "excelDataExport")
                button.click()
            except Exception:
                try:
                    button = self.driver.find_element(By.ID, "download")
                    button.click()
                except Exception as e:
                    print(f"An error occurred while clicking the button: {e}")

    # Function to get the system's default download folder
    def get_default_download_folder(self):
        if os.name == 'nt':  # For Windows
            download_folder = str(Path.home() / "Downloads")
        else:  # For macOS/Linux
            download_folder = str(Path.home() / "Downloads")
        return download_folder

    # Function to wait for download to finish
    def wait_for_download(self, timeout=30):
        end_time = time.time() + timeout
        while True:
            if any(file.endswith(".crdownload") for file in os.listdir(self.download_folder)):
                time.sleep(1)
            else:
                break
            if time.time() > end_time:
                raise TimeoutError("Download did not finish in the allocated time.")

    # Function to move downloaded files to the target folder
    def move_downloaded_files(self):
        files_moved = []
        for file in os.listdir(self.download_folder):
            if file.endswith('.csv') or file.endswith('.xlsx'):
                full_file_name = os.path.join(self.download_folder, file)
                target_path = os.path.join(self.target_folder, file)
                shutil.move(full_file_name, target_path)
                files_moved.append(target_path)
        return files_moved

    # Function to delete files after processing
    def delete_files(self, files):
        for file in files:
            try:
                os.remove(file)
                print(f"Deleted file: {file}")
            except OSError as e:
                print(f"Error deleting file {file}: {e}")

    # Main function to download reports and load into DataFrames
    def download_and_process_reports(self):
        # Get previous month range
        start_date, end_date = self.get_previous_month_range()

        # URLs
        admission_base_url = "https://aloraplus.com/Report/RptAdmDetail?dateFrom={}&dateThrough={}&rdOffice=0&officesList_length=10&patient=&patientName=&patientAll=true&rdPayer=1&payersList_length=10&pay-42=42&groupBy=N&includeAdmBefore=true&showFrequency=true&caregiver=&caregiverName=&caseManagerAll=true&physician=&physicianName=&physicianAll=true&evacLevel=&evacLevelName=&evacLevelAll=true&region=&regionName=&regionAll=true&careTeamMember=&careTeamMemberAll=true&admTrack1Num=&displayTelephony=true&cdIdsArray=13%2C41%2C9%2C42&officeIdsArray=&admTrack1=&admTrack2=&admTrack3=&admTrack4=&admTrack5=&admTrack6=&optSettingName=&setPublic=1&DataField=&Relational=&UserText=&optSettings_length=10&reportType=AdmDetail&optSettings=%5B%5D"
        visit_list_base_url = "https://aloraplus.com/Report/RptVisitList?dateFrom={}&dateThrough={}}&patient=&patientAll=true&caregiverVisit=&caregiverAll=true&visitStatus=All&groupBy=P&sortBy=VD&office=&officeAll=true&billingCode=&billingCodeAll=true&allDisciplines=true&disciplineNameOrType=N&discipline=-1&payer=&payerAll=true&payrollStatus=A&includeAmount=None&billingStatus=A&billable=A&patientRegion=&patientRegionAll=true&patientName=&officeName=&caregiverName=&billingCodeName=&payerName=&regionName=&disId=&disText=&cdIdsArray=undefined"
        physical_therapy_docs_base_url = "https://aloraplus.com/PT/Summary?id={}"
        occupational_therapy_docs_base_url = "https://aloraplus.com/OT/Summary?id={}"
        speech_therapy_docs_base_url = "https://aloraplus.com/ST/STSummary?id={}"
        medical_social_worker_docs_base_url = "https://aloraplus.com/MedicalSocialWorker/Summary?id={}"

        updated_admission_url = admission_base_url.format(start_date, end_date)
        updated_visit_list_url = visit_list_base_url.format(start_date, end_date)
        report_list = [updated_admission_url, updated_visit_list_url]

        # Login to website
        self.driver.get("https://aloraplus.com/Report/Index")
        username_field = self.driver.find_element(By.ID, 'UserName')
        password_field = self.driver.find_element(By.ID, 'Password')
        submit_button = self.driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]')
        username_field.send_keys("jackermann@abilityhomehealthservices.com")
        password_field.send_keys("abs$06162003$J")
        submit_button.click()

        # Download each report and wait for completion
        for url in report_list:
            self.driver.get(url)
            self.click_button_with_fallback()
            self.wait_for_download()  # Wait until the file finishes downloading

        # Move downloaded files to the target folder
        downloaded_files = self.move_downloaded_files()

        # Load each file into a DataFrame and delete after uploading
        for file in downloaded_files:
            if file.endswith('.csv'):
                df = pd.read_csv(file)
            elif file.endswith('.xlsx'):
                df = pd.read_excel(file)
            self.dfs.append(df)  # Add DataFrame to the list

        # Delete files after they are loaded
        self.delete_files(downloaded_files)

    # Function to close the driver and clean up
    def close(self):
        self.driver.quit()

# Usage example:
if __name__ == "__main__":
    downloader = ReportDownloader()
    downloader.download_and_process_reports()  # Download and process reports
    for idx, df in enumerate(downloader.dfs):
        print(f"DataFrame {idx + 1}:")
        print(df.head())  # Display first few rows of each DataFrame
    downloader.close()  # Clean up and close the browser
