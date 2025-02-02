import pandas as pd
import numpy as np
import os

def client_Invoice(file1, output_dir="output"):
    """
    Processes the input Excel file to generate a client invoice file.

    Parameters:
        file1 (str): Path to the input Excel file.
        output_dir (str): Directory to save the output file.

    Returns:
        str: Path to the generated client invoice Excel file.
        list: List of error or status messages.
    """
    # Output file name and path
    file2 = os.path.join(output_dir, "Client_Invoice.xls")
    msg = []

    # Columns to keep in the output file
    selected_columns = ['VoiceFile Name', 'Document Name', 'Doctor', 'DOS', 'Line Count', 'Remarks']

    try:
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)

        # Read all sheets into a dictionary
        all_sheets = pd.read_excel(file1, sheet_name=None)

        # Create a new Excel writer
        with pd.ExcelWriter(file2, engine='xlsxwriter') as writer:
            for sheet_name, df in all_sheets.items():
                # Check if 'Line Count' column exists in the DataFrame
                if 'Line Count' in df.columns:
                    # Add the sum of 'Line Count' to the sheet
                    # Safely add a 'Total' row by creating a new DataFrame row with the correct structure
                    total_row = pd.DataFrame([{'Line Count': df['Line Count'].sum()}], index=['Total'])
                    df = pd.concat([df, total_row], axis=0)


                    # Ensure all selected columns are present
                    if all(col in df.columns for col in selected_columns):
                        try:
                            # Convert 'DOS' to proper date format
                            df['DOS'] = pd.to_datetime(df['DOS'], errors='coerce').dt.strftime('%m/%d/%Y')
                        except Exception as e:
                            msg.append(f"Error formatting 'DOS' in sheet '{sheet_name}': {e}")

                        # Select only the required columns
                        selected_df = df[selected_columns]
                        selected_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        msg.append(f"Skipping sheet '{sheet_name}' due to missing columns.")
                else:
                    msg.append(f"Skipping sheet '{sheet_name}' due to missing 'Line Count' column.")
    except Exception as e:
        msg.append(f"An error occurred: {e}")
        return None, msg

    return file2, msg