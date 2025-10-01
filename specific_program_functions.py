import os
from datetime import date, timedelta
from openpyxl import Workbook


def get_previous_month_path():
    """Return and create a folder path for previous month in the required format."""

    # Polish month names mapping
    MONTHS_PL = {
        1: "Styczeń", 2: "Luty", 3: "Marzec", 4: "Kwiecień",
        5: "Maj", 6: "Czerwiec", 7: "Lipiec", 8: "Sierpień",
        9: "Wrzesień", 10: "Październik", 11: "Listopad", 12: "Grudzień"
    }

    # Determine previous month and year
    today = date.today()
    first_this_month = today.replace(day=1)
    last_prev_month = first_this_month - timedelta(days=1)
    prev_month = last_prev_month.month
    prev_year = last_prev_month.year

    # Build folder structure
    username = os.getlogin()
    base_path = fr"C:\Users\{username}\OneDrive - Roto Frank DST\General\02_Raporty\11_Odpad_WOD\{prev_year}"
    folder_name_base = f"{prev_month:02d}_odpad_{prev_year}_{MONTHS_PL[prev_month]}"

    # Ensure base path exists
    if not os.path.exists(base_path):
        os.makedirs(base_path)

    # Determine full folder path (with index if necessary)
    folder_path = os.path.join(base_path, folder_name_base)
    idx = 0

    while os.path.exists(folder_path):
        idx += 1
        folder_path = os.path.join(base_path, f"{folder_name_base}_{idx:02d}")

    # Create final folder
    os.makedirs(folder_path)
    return folder_path, f"{folder_name_base}_{idx:02d}"


def create_dummy_excel_file_listwa(files_path, file_name):
    # Ensure directory exists
    os.makedirs(files_path, exist_ok=True)

    # Full path for the file
    file_path = os.path.join(files_path, file_name + ".XLSX")

    # Create a new workbook and select active sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Define headers
    headers = [
        "Zlecenie",
        "Nr materiału",
        "Krótki tekst materiału",
        "Ilość dostarczona (GMEIN)",
        "Jednostka miary (=GMEIN)"
    ]

    # Add headers to first row
    ws.append(headers)

    # Add a row with zeros (or empty string for text columns if you prefer)
    zero_row = [0, 0, 0, 0, 0]
    ws.append(zero_row)

    # Save file
    wb.save(file_path)
    print(f"Excel file created: {file_path}")

