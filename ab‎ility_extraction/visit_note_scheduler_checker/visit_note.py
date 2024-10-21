from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import os
import shutil
import pdfplumber
from datetime import datetime, timedelta
import pandas as pd

def get_previous_month_range():
    today = datetime.now()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    first_day_of_previous_month = last_day_of_previous_month.replace(day=1)
    return first_day_of_previous_month.strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d')

def click_button_with_fallback(driver):
    try:
        button = driver.find_element(By.ID, "csvDataExport")
        button.click()
    except Exception:
        try:
            button = driver.find_element(By.ID, "excelDataExport")
            button.click()
        except Exception:
            try:
                button = driver.find_element(By.ID, "download")
                button.click()
            except Exception as e:
                print(f"An error occurred while clicking the button: {e}")
                
start_date, end_date = get_previous_month_range()

admission_base_url = "https://aloraplus.com/Report/RptAdmDetail?dateFrom={}&dateThrough={}&rdOffice=0&officesList_length=10&patient=&patientName=&patientAll=true&rdPayer=1&payersList_length=10&pay-42=42&groupBy=N&includeAdmBefore=true&showFrequency=true&caregiver=&caregiverName=&caseManagerAll=true&physician=&physicianName=&physicianAll=true&evacLevel=&evacLevelName=&evacLevelAll=true&region=&regionName=&regionAll=true&careTeamMember=&careTeamMemberAll=true&admTrack1Num=&displayTelephony=true&cdIdsArray=13%2C41%2C9%2C42&officeIdsArray=&admTrack1=&admTrack2=&admTrack3=&admTrack4=&admTrack5=&admTrack6=&optSettingName=&setPublic=1&DataField=&Relational=&UserText=&optSettings_length=10&reportType=AdmDetail&optSettings=%5B%5D"
visit_list_base_url = "https://aloraplus.com/Report/RptVisitList?dateFrom={}&dateThrough={}}&patient=&patientAll=true&caregiverVisit=&caregiverAll=true&visitStatus=All&groupBy=P&sortBy=VD&office=&officeAll=true&billingCode=&billingCodeAll=true&allDisciplines=true&disciplineNameOrType=N&discipline=-1&payer=&payerAll=true&payrollStatus=A&includeAmount=None&billingStatus=A&billable=A&patientRegion=&patientRegionAll=true&patientName=&officeName=&caregiverName=&billingCodeName=&payerName=&regionName=&disId=&disText=&cdIdsArray=undefined"
physical_therapy_docs_base_url = "https://aloraplus.com/PT/Summary?id={}"
occupational_therapy_docs_base_url = "https://aloraplus.com/OT/Summary?id={}"
speech_therapy_docs_base_url = "https://aloraplus.com/ST/STSummary?id={}"
medical_social_worker_docs_base_url = "https://aloraplus.com/MedicalSocialWorker/Summary?id={}"

updated_admission_url = admission_base_url.format(start_date, end_date)
updated_visit_list_url = visit_list_base_url.format(start_date, end_date)

report_list = [updated_admission_url, updated_visit_list_url]


