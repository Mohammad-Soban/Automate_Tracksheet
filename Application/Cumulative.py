import os
import openpyxl
import pandas as pd
import numpy as np
import csv
from xlutils.copy import copy

def middle_for_cum(file1, output_dir='output'):
    middle_2 = os.path.join(output_dir, "Middle_2.xls")
    msg = []

    try:
        selected_columns = ['VoiceFile Name', 'Document Name', 'Doctor', 'Line Count']

        # Read all sheets into a dictionary
        all_sheets = pd.read_excel(file1, sheet_name=None)

        # Create a new Excel writer
        with pd.ExcelWriter(middle_2, engine='xlsxwriter') as writer:
            for sheet_name, df in all_sheets.items():
                # Check if 'Line Count' column exists in the DataFrame
                if 'Line Count' in df.columns:
                    # Find the sum of the column 'Line Count' for each sheet and append it to column 'Line Count' in the sheet at the end
                    df.loc['Total'] = pd.Series(df['Line Count'].sum(), index=['Line Count'])

                    # Ensure all selected columns are present in the DataFrame
                    if all(col in df.columns for col in selected_columns):
                        selected_df = df[selected_columns]

                        # Write the selected DataFrame to the new workbook
                        selected_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        msg.append(f"Skipping sheet '{sheet_name}' due to missing columns.")
                else:
                    msg.append(f"Skipping sheet '{sheet_name}' due to missing 'Line Count' column.")
    except Exception as e:
        msg.append(f"An error occurred: {e}")
        return None, msg

    return middle_2, msg

def Cumulative(file1, output_dir="output"):
    middle_2, msg = middle_for_cum(file1, output_dir)
    if not middle_2:
        return None, msg

    source_workbook_paths = [
        "./output/Client_Invoice.xls",
        "./output/Middle_1.xls",
        "./output/Middle_2.xls"
    ]

    cumulative = os.path.join(output_dir, "Cumulative.xls")
    msg = []

    # Create an empty DataFrame to store the results
    summary_df = pd.DataFrame(columns=['Sheet Name', 'Total Line Count'])

    try:
        # Iterate through each source workbook
        for source_workbook_path in source_workbook_paths:
            # Check if the file exists
            if not os.path.exists(source_workbook_path):
                msg.append(f"File not found: {source_workbook_path}")
                continue

            # Read the source workbook
            xl = pd.ExcelFile(source_workbook_path)

            # Iterate through each sheet in the source workbook
            for sheet_name in xl.sheet_names:
                # Read the sheet
                df = pd.read_excel(source_workbook_path, sheet_name)

                # Check if the DataFrame has any rows
                if not df.empty:
                    # Extract the last value of the "Line Count" column
                    total_Line_Count = df['Line Count'].iloc[-1]

                    # Append the result to the summary DataFrame
                    summary_df = pd.concat([
                        summary_df,
                        pd.DataFrame({'Sheet Name': [sheet_name], 'Total Line Count': [total_Line_Count]})
                    ], ignore_index=True)
                else:
                    msg.append(f"Skipping sheet '{sheet_name}' in '{source_workbook_path}' due to empty DataFrame.")

        # Create a new workbook and write the summary DataFrame to a sheet
        with pd.ExcelWriter(cumulative, engine='xlsxwriter') as writer:
            summary_df.to_excel(writer, sheet_name='Summary', index=False, header=True)
    except Exception as e:
        msg.append(f"An error occurred: {e}")
        return None, msg

    return cumulative, msg