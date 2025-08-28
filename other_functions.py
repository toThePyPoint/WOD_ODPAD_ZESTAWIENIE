import os

import pandas as pd
import pyperclip
import io
import win32com.client


def copy_df_column_to_clipboard(df, column_name):
    """
    Copies the specified column from a pandas DataFrame to the clipboard using pyperclip.

    Args:
        df (pd.DataFrame): The pandas DataFrame containing the data.
        column_name (str): The name of the column to copy.

    Returns:
        bool: True if the data was copied successfully, False otherwise.
    """
    try:
        # Check if the column exists in the DataFrame
        if column_name not in df.columns:
            print(f"Error: Column '{column_name}' not found in DataFrame.")
            return False

        # # Select the column and convert it to a string format
        # column_data = df[column_name].astype(str)
        # column_string = column_data.to_string(header=True, index=False)

        # Select the column
        column_data = df[column_name]

        # Convert to string with tab separation
        output = io.StringIO()
        column_data.to_csv(output, sep='\t', header=False, index=False)
        column_string = output.getvalue()
        output.close()

        # Copy the string to the clipboard using pyperclip
        pyperclip.copy(column_string)
        return True
    except Exception as e:
        print(f"An error occurred: {e}")
        return False


def close_excel_file(file_name):
    try:
        # Connect to the running Excel application
        excel = win32com.client.Dispatch("Excel.Application")

        # Iterate through open workbooks
        for workbook in excel.Workbooks:
            if workbook.FullName.lower().endswith(file_name.lower()):  # Match the file name
                workbook.Save()  # Ensure the file is saved
                workbook.Close()  # Close the workbook
                print(f"{file_name} has been saved and closed.")
                # Quit Excel if no other workbooks are open
                # if excel.Workbooks.Count == 0:
                #     excel.Quit()
                return True

        print(f"{file_name} not found in open Excel instances.")
        return False

    except Exception as e:
        print(f"An error occurred: {e}")
