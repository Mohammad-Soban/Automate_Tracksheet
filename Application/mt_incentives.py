import pandas as pd
import xlwt
import openpyxl
import numpy as np
import csv
import os

def MT_incentives(file1, output_dir="output"):
    """
    Processes the input Excel file to generate an MT Invoice file.
    
    Parameters:
        file1 (str): Path to the input Excel file.
        output_dir (str): Path to the output directory

    Returns:
        list: Paths to the generated files.
        list: List of error or status messages.
    """
    middle_1 = os.path.join(output_dir, "Middle_1.xls")
    mt_incentives = os.path.join(output_dir, "MT_incentives.xls")
    msg = []

    try:
        all_sheets = pd.read_excel(file1, sheet_name=None)
    except Exception as e:
        msg.append(f"Error reading Excel file: {e}")
        return None, msg

    unique_values_2 = set()

    try:
        for sheet_name, df in all_sheets.items():
            if 'Transcribed By' in df.columns:
                df.dropna(subset=['Transcribed By'], inplace=True)
                unique_values_2.update(df['Transcribed By'].str[:2].unique())
            else:
                msg.append(f"Skipping sheet '{sheet_name}' due to missing 'Transcribed By' column.")
    except AttributeError as e:
        msg.append(f"Error processing sheets: {e}")
        return None, msg

    columns = ["VoiceFile Name", "Document Name", "DOB", "Doctor", "Date on FTP", "Download Date", "DOS", "Transcribe Date", "Transcribed By", "Line Count"]
    created_sheets = set()

    try:
        with pd.ExcelWriter(middle_1, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets.items():
                if 'Transcribed By' in df.columns:
                    for value in unique_values_2:
                        df['Transcribed By'] = df['Transcribed By'].astype(str)
                        filtered_rows = df[df['Transcribed By'].str.startswith(value)]
                        sheet_name_to_create = f'{sheet_name}_{value}'
                        filtered_rows.to_excel(writer, sheet_name=sheet_name_to_create, index=False)
                        created_sheets.add(sheet_name_to_create)
                else:
                    msg.append(f"Skipping sheet '{sheet_name}' due to missing 'Transcribed By' column.")
    except Exception as e:
        msg.append(f"Error writing to middle_1 file: {e}")
        return None, msg

    try:
        with pd.ExcelWriter(mt_incentives, engine='openpyxl') as writer:
            for suffix in set(sheet_name[-2:] for sheet_name in created_sheets):
                merged_df = pd.concat([pd.read_excel(middle_1, sheet_name=sheet_name)[columns] for sheet_name in created_sheets if sheet_name.endswith(suffix)])
                for date_column in ["DOB", "Date on FTP", "Download Date", "DOS", "Transcribe Date"]:
                    merged_df[date_column] = pd.to_datetime(merged_df[date_column], errors='coerce').dt.strftime('%m/%d/%Y')
                total_line_count = merged_df["Line Count"].sum()
                total_row = pd.DataFrame({"Line Count": [total_line_count]})
                merged_df = pd.concat([merged_df, total_row])
                merged_df.to_excel(writer, sheet_name=suffix, index=False, header=columns)
    except Exception as e:
        msg.append(f"Error writing to MT_incentives file: {e}")
        return None, msg

    return [middle_1, mt_incentives], msg