import pandas as pd
import datetime
from datetime import date
import numpy as np
import os


def plant_val(row):
    if row["Last name, first name"] is not None:
        return 2115

def date_col(row, work_date):
    if row["Last name, first name"] is not None:
        return f"({work_date})"

def conf_text(row):
    if row["Last name, first name"] is not None:
        temp_name = row["Last name, first name"]
        temp_name = temp_name.replace(", ", ";")
        date = row["Date Worked"]
        if row["Sub"] == "1" and row["Nights"] == "Yes":
            return f"{temp_name} {date} SUB;NS"
        elif row["Sub"] == "1":
            return f"{temp_name} {date} SUB"
        elif row["Nights"] == "Yes":
            return f"{temp_name} {date} NS"
        else:
            return f"{temp_name} {date}"


def nova_upload(file_path):
    df_ts = pd.read_excel(file_path, engine='openpyxl', sheet_name='Timesheet', header=7, usecols='A:F,K,M', dtype='str').dropna(how='all')
    df_emp_list = pd.read_excel(file_path, engine = 'openpyxl', sheet_name = "Employee List", usecols="B,C,F,G", dtype='str').dropna(how='all')
    df_date = pd.read_excel(file_path, sheet_name='Timesheet', engine='openpyxl', usecols='B', header=1)


    temp_col = df_date.columns
    work_date = temp_col[0]
    work_date = datetime.datetime.strftime(work_date, "%b %d,%Y")
 
    df_nova_upload = df_ts.merge(df_emp_list, left_on="Last name, first name", right_on="Last name, first name", how='left')

    df_nova_upload['Plant'] = df_ts.apply(lambda row: plant_val(row), axis=1)



    df_nova_upload['Date Worked'] = df_nova_upload.apply(lambda row: date_col(row, work_date), axis=1)



    df_nova_upload['Confirmation Text'] = df_nova_upload.apply(lambda row: conf_text(row), axis=1)



    df_nova_upload['Posting Date'] = date.today().strftime("%Y%m%d")
    
    df_nova_upload = df_nova_upload[['Work order #', 'Op', 'Sub-op', 'WorkCentre', 'Plant', 'Total Hours', 'OT Activity', 'Conf', 'Confirmation Text', 'Posting Date', 'Quinn #', 'Last name, first name', 'Date Worked', 'Sub', 'Nights']]

    col_headers = ['WO',	'Operation',	'Sub-Op',	'Work Centre',	'Plant',	'Hours',	'Activity',	'Final CNF',	'Confirmation Text',	'Posting Date',	'Employee',	'Name',	'Date Worked',	'Sub',	'Nights']
    df_nova_upload.columns = col_headers

    df_nova_upload['Sub'] = df_nova_upload['Sub'].replace("1", "Sub")
    df_nova_upload['Sub'] = df_nova_upload['Sub'].replace("0", "")

    df_nova_upload['Nights'] = df_nova_upload['Nights'].replace("Yes", "NS")

    df_nova_upload.replace(np.NaN, "", inplace=True)

    nova_upload_csv = os.path.join(os.getcwd(), "nova_upload.csv")
    df_nova_upload.to_csv(nova_upload_csv, index=False)
    
    # nova_csv_bin_path = open(nova_upload_csv, 'rb', buffering=0)
    # nova_csv_binary = nova_csv_bin_path.read()
    # nova_csv_bin_path.close()
    # os.remove(nova_upload_csv)

    # return nova_csv_binary

    
file_path = r"C:\Users\divyal\Desktop\projects\NOVA\ta_ts\23-E1TA-20210709-10.xlsm"
nova_upload(file_path)