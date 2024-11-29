import pandas as pd
import numpy as np
import os
from constants import region_list

# Read Checklist Audit Excel
def read_ChecklistAudit(directory):
    files = [file for file in os.listdir(directory) if file.startswith('Checklist') and file.endswith('.xlsx')]
    file_path = os.path.join(directory, files[0])

    df = pd.read_excel(file_path)

    # Remove triage
    df = df[df['TypeOfAssessment'] != 'Triage']

    return df

# Read Hand Hygiene Audit Excel
def read_HandHygieneAudit(directory):
    files = [file for file in os.listdir(directory) if file.startswith('Hand Hygiene') and file.endswith('.xlsx')]

    df_array = []
    
    if files:
        for file in files:
            file_path = os.path.join(directory, file)
            df = pd.read_excel(file_path)

            df = df[df['Status'] == 'Closed']

            df_array.append(df)

        df_HH = pd.concat(df_array)

        # df_HH['AssessmentDate'] = pd.to_datetime(df_HH['AssessmentDate'], format='ISO8601', errors='coerce')

        return df_HH
    
# Read Pharmacy and Medication Storage Audit Excel
def read_PMSAudit(directory):
    files = [file for file in os.listdir(directory) if file.startswith('Pharmacy and medication') and file.endswith('.xlsx')]
    file_path = os.path.join(directory, files[0])

    df = pd.read_excel(file_path)

    return df

# Read SamplingAudit Excel
def read_SamplingAudit(directory):
    files = [file for file in os.listdir(directory) if file.startswith('Sampling') and file.endswith('.xlsx')]
    file_path = os.path.join(directory, files[0])

    df = pd.read_excel(file_path)

    return df

area = {
    'Hygiene and Infection control' : 'Hygiene and Infection Control',
    'Hand Higiene' : 'Hygiene and Infection Control',
    'Donning and doffing PPE' : 'Hygiene and Infection Control',

    'Needle Taping' : 'Vascular Access', 
    'CVC connections' : 'Vascular Access', 
    'CVC Care': 'Vascular Access', 
    'AVF AVG Care' :  'Vascular Access',

    'Healthcare records' : 'Documentation',
    'Dialysis Prescriptions - delivery' : 'Documentation', 
    'Logs': 'Documentation',

    'Medication preparation and administration' : 'Medication', 
    'Audit pharmacy' : 'Medication',

    'Reuse' : 'Reuse'
}



# Load all functions to read audit excels
def loadAll(directories):
    checklist_df = read_ChecklistAudit(directories[0])
    hh_df = read_HandHygieneAudit(directories[1])
    pms_df = read_PMSAudit(directories[2])
    sampling_df = read_SamplingAudit(directories[3])

    # Combine all df
    dataLength_array = [
        len(checklist_df),
        len(hh_df),
        len(pms_df),
        len(sampling_df)
    ]

    df = pd.concat([checklist_df, hh_df, pms_df, sampling_df])

    # Add Region
    df['Region'] = df['Clinic'].map(region_list)

    # Add Month & Quarter
    # Convert to timezone-naive and add the Quarter column
    # Convert all datetime values to timezone-naive
    df['AssessmentDate'] = pd.to_datetime(df['AssessmentDate'], errors='coerce', utc=True).dt.tz_localize(None)
    df['Quarter'] = df['AssessmentDate'].dt.to_period('Q')

    # Area column
    df['Area'] = df['TypeOfAssessment'].map(area)

    # Apply logic to fill 'External Audit' column
    auditTypeCSS = ['Jeyasutha Sukumar', 'Mohana Regunathan', 'Nurul Hawa Salim', 'Thomasraj Soosainathan']
    df['External_Audit'] = df['External_Audit'].fillna(
        df['Registered by'].apply(lambda x: 'External Audit' if x in auditTypeCSS else 'Internal Audit')
    )

    # Rename External_Audit as Category
    df = df.rename({'External_Audit': 'Category'}, axis=1)

    print(f"Sum lenght: {sum(dataLength_array)}, Final size: {len(df)}, diff: {len(df) - sum(dataLength_array)}")

    return df

# Function to create table for external audit
def externalAuditTable(df):
    df['Done'] = 'Yes'

    df['Quarter'] = df['Quarter'].astype(str)
    df = df[df['Quarter'].str.contains('2024')]

    df = df[df['Category'] == 'External Audit']

    #Pivot the table
    results = (
        df.pivot_table(
            index=["Region", "Clinic"],  # Rows: Clinics
            columns=["TypeOfAssessment", "Quarter"],  # Columns: Type of Assessment and Quarter
            values="Done",  # Values: "Yes"/"No" indicator
            aggfunc=lambda x: "Yes" if len(x) > 0 else "No"  # Aggregate to check for presence
        )
        .fillna("No")  # Fill missing values with "No"
    )

    results = results.reset_index()

    return results






        
