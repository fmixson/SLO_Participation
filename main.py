import pandas as pd
import openpyxl
from openpyxl import load_workbook
import lxml
from configparser import ConfigParser
import time
import os, sys

# Import SLO Participation report
faculty_participation_df = pd.read_csv(
    'C:/Users/fmixson/Desktop/SLOs/Fall 23 SLO Participation 03_04_2024.csv', encoding='latin-1')

pd.set_option('display.max_columns', None)
faculty_participation_df = faculty_participation_df.fillna(0)
course_sections_df = faculty_participation_df[faculty_participation_df['Course or Section'].str.contains('Section', na=False)]
course_sections_df[['Course', 'Class#','Delete', 'To Be Deleted']] = course_sections_df['Course or Section'].str.split(' ', expand=True)
course_sections_df[['Completed', 'Of', 'Total Assessments']] = course_sections_df['Completed Assessments'].str.split(' ', expand=True)
# df[['First Name', 'Last Name']] = df['Name'].str.split(' ', expand=True)
course_sections_df = course_sections_df[course_sections_df['Class#'] != 'Totals']



# Import merged worksheets
fall_schedule_df = pd.read_csv(
    'C:/Users/fmixson/Desktop/SLOs/Fall23_dataframe.csv', encoding='latin-1')
pd.set_option('display.max_columns', None)

# Merge the two worksheets
print(fall_schedule_df.dtypes, course_sections_df.dtypes)
fall_schedule_df['Class#'] = fall_schedule_df['Class#'].astype('str')
# course_sections_df['Completed'] = course_sections_df['Completed'].astype('float')
# course_sections_df['Total Assessments'] = course_sections_df['Total Assessments'].astype('float')

merged_df = pd.merge(fall_schedule_df, course_sections_df, on=['Class#'])

merged_df = merged_df[['Combined','Division', 'Dept','Course_x', 'Class#', 'Session', 'Modality', 'Component', 'Start', 'End',
                               'Days', 'Instructor', 'Room', 'Completed', 'Total Assessments']]
merged_df['Completed'] = merged_df['Completed'].astype(int)
print(merged_df.dtypes)
merged_df['Total Assessments'] = merged_df['Total Assessments'].astype(int)
# df[["a", "b"]] = df[["a", "b"]].apply(pd.to_numeric)
merged_df.to_excel('Merged_dataframes.xlsx')


# Filter out labs