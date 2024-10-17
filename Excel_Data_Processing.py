'''
Author: Tom Diederen
Summary: this script takes an Excel file as input for a Pandas dataframe. Processes data and outputs it to a (different) Excel file.
    The input .xlsx can be an export from a database or other program, the output can be a consolidated spreadsheet with insights from different sources.
'''

import pyautogui
import time
import datetime
import openpyxl
import pandas as pd
from pathlib import Path
import os
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import keyboard_shortcuts as ks

today = datetime.date.today()
year, week, day_of_week = today.isocalendar()
input_file = f"Analysis_input_wk{week}.{day_of_week}_{year}.xlsx"   # Excel file that holds the input data.                    
export_path = f"Analysis_output_wk{week}.{day_of_week}_{year}"      # Excel file that will hold the output data.
output_sheet = 'Output_data'                                        # Sheet name withint the output Excel file.
duplicate_input_files = 1                                           # Will count up if duplicate input file names exists. If "file_name", already exists, "file_name_1" is saved instead.
duplicate_output_files = 1                                          # Same for output files.
pause = 0.5                                                         # Pause before and after keystrokes, in seconds. Needed so the system can keep up with fast keyboard and mouse input.
df_payload = None                                                   # Dataframe that holds input data and can be analyzed / processed to provide desired output.

# Example sequence of actions to gather input data from a database or other program. Can be skipped if file is already available. Writes to input_file.
def gather_input_data():
    global input_file
    global duplicate_input_files
    renamed_input_file = input_file
    path = Path(f"{input_file}.xlsx")

    # Example clicks to navigate to export option of SW used:
    pyautogui.click(x=10, y=102)
    time.sleep(1)

    pyautogui.click(x=10, y=200)

    # Check for duplicate file name
    while path.exists():
        if duplicate_input_files == 1:  # First duplicate, add "_1" to file name.
            path = path.with_name(f"{path.stem}_{duplicate_input_files}{path.suffix}")
        else:
            stem = path.stem
            versionless_stem = stem[:-2]
            path = path.with_name(f"{versionless_stem}_{duplicate_input_files}{path.suffix}")
        duplicate_input_files += 1
    renamed_input_file = str(path)
    input_file = renamed_input_file
    
    # Save input file
    pyautogui.write(f"{renamed_input_file}")
    time.sleep(1)
    pyautogui.press("enter")

# Read xls -> dataframe, perform analysis
def analyze():
    # df_payload = pd.read_excel(f"{os.path.abspath(input_file)}")    # Use after automatic data import (with gather_input_data() )
    df_payload = pd.read_excel("Financial_Sample.xlsx")               # Use for manually selecting data import source. Example: Financial sample data downloaded from: https://learn.microsoft.com/en-us/power-bi/create-reports/sample-financial-download#download-it-directly
    gross_per_country = df_payload.groupby('Country')['Gross Sales'].sum().reset_index() # Sample analysis: calculate gross sales per country
    gross_per_country['Gross Sales'] = gross_per_country['Gross Sales'] / 1_000_000  # Round to millions to make plot easier to read
    def millions_formatter(x, pos):
        return f'${x:.3f}M'

    # Create bar plot of results
    sns.barplot(data=gross_per_country, x='Country', y='Gross Sales', hue='Country')  

    # Add labels and title
    plt.title('Gross Sales by Country')
    plt.xlabel('Country')
    plt.ylabel('Sales (in Millions USD)')
    plt.gca().yaxis.set_major_formatter(FuncFormatter(millions_formatter))
    plt.show()

def export_to_xlsx():
    book = openpyxl.load_workbook(export_path)
    sheet = book[output_sheet]
    with pd.ExcelWriter(export_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df_payload.to_excel(writer, sheet_name=output_sheet, startrow=1, startcol=writer.sheets[output_sheet].max_column + 1, index=False, header=True)
        print(f"Data exported to: {export_path}")

'''
Script
'''
gather_input_data()
analyze()
export_to_xlsx()



