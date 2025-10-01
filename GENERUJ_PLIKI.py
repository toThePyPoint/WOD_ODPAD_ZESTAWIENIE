import os
import sys
import ctypes
import traceback
import pandas as pd

from sap_connection import get_last_session
from sap_functions import open_one_transaction, simple_load_variant
from gui_manager import adjust_previous_month_dates_by_user, show_message
from specific_program_functions import get_previous_month_path, create_dummy_excel_file_listwa
from sap_functions import (export_data_to_file_MB51, export_data_to_file_COHV, paste_production_orders_and_load_variant,
                           export_data_to_file_ZEK1, zek1_change_layout)
from other_functions import copy_df_column_to_clipboard
from sap_transactions import cohv_go_back_change_layout_list_type_and_order_type


MB51_VARIANT_NAME = "REP_LU_WOD_EFF"
MB51_VARIANT_NAME_DREWNO = "REP_LU_WOOD001"
MB51_VARIANT_NAME_KANTOWKA = "REP_LU_WOOD002"
COHV_VARIANT_NAME_LISTWA = "REP_LU_WOOD001"
ZEK1_VARIANT_NAME = "REP_LU_WOOD001"


if __name__ == "__main__":
    # Hide console window
    if sys.platform == "win32":
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        hWnd = kernel32.GetConsoleWindow()
        if hWnd:
            user32.ShowWindow(hWnd, 6)  # 6 = Minimize

    try:
        files_path, folder_name = get_previous_month_path()  # get path and folder name
        first_date, last_date, zek1_range_left, zek1_range_right = adjust_previous_month_dates_by_user(message=f"Nazwa katalogu: {folder_name}")  # get the dates

        # get session, transaction and number of window in SAP
        sess1, tr1, nu1 = get_last_session(max_num_of_sessions=6)

        open_one_transaction(session=sess1, transaction_name='MB51')
        simple_load_variant(obj_sess=sess1, variant_name=MB51_VARIANT_NAME, open_only=True)

        # Adjust dates
        sess1.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = first_date
        sess1.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = last_date
        sess1.findById("wnd[0]").sendVKey(8)

        export_data_to_file_MB51(sess1, files_path, "mb51.XLSX")

        mb51_df = pd.read_excel(os.path.join(files_path, "mb51.XLSX"), dtype={"Zlecenie": "str"})
        mb51_df.dropna(subset=['Zlecenie'], inplace=True)
        copy_df_column_to_clipboard(mb51_df, 'Zlecenie')

        open_one_transaction(session=sess1, transaction_name="COHV")

        sess1.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOR000"
        sess1.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/ODPAD3"
        sess1.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_LOEKZ").selected = True
        sess1.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").text = "2101"
        sess1.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_AUART-LOW").text = "RO09"
        paste_production_orders_and_load_variant(sess1)
        export_data_to_file_COHV(sess1, files_path, "potwierdzenia.XLSX")

        cohv_go_back_change_layout_list_type_and_order_type(session=sess1, list_type="PPIOM000", layout="000000000001", ord_type='RO09')
        export_data_to_file_COHV(sess1, files_path, "składniki.XLSX")

        cohv_go_back_change_layout_list_type_and_order_type(session=sess1, list_type="PPIOH000", layout="/MARCINW", ord_type='RO07')
        export_data_to_file_COHV(sess1, files_path, "odzysk_ro07.XLSX")

        mb51_migo_booking_variants = [MB51_VARIANT_NAME_DREWNO, MB51_VARIANT_NAME_KANTOWKA]
        mb51_migo_booking_file_names = ["mb51_drewno_odpad_migo.XLSX", "mb51_kantowka_odpad_migo.XLSX"]

        for variant, f_name in zip(mb51_migo_booking_variants, mb51_migo_booking_file_names):
            open_one_transaction(session=sess1, transaction_name='MB51')
            simple_load_variant(obj_sess=sess1, variant_name=variant, open_only=True)

            # Adjust dates
            sess1.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = first_date
            sess1.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = last_date
            sess1.findById("wnd[0]").sendVKey(8)  # Load variant
            export_data_to_file_MB51(sess1, files_path, f_name)

        open_one_transaction(session=sess1, transaction_name="COHV")
        simple_load_variant(obj_sess=sess1, variant_name=COHV_VARIANT_NAME_LISTWA, open_only=True)

        # Adjust dates
        sess1.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").text = first_date
        sess1.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").text = last_date
        sess1.findById("wnd[0]").sendVKey(8)
        try:
            # If there is no data
            sess1.findById("wnd[1]/tbar[0]/btn[0]").press()
            create_dummy_excel_file_listwa(files_path, "listwa.XLSX")
            missing_data = True
            print("Missing data: ", missing_data)
        except:
            missing_data = False
            print("ERROR! Missing data: ", missing_data)
        if not missing_data:
            export_data_to_file_COHV(sess1, files_path, "listwa.XLSX")

        open_one_transaction(session=sess1, transaction_name="ZEK1")
        simple_load_variant(obj_sess=sess1, variant_name=ZEK1_VARIANT_NAME, open_only=True)
        # Adjust range
        sess1.findById("wnd[0]/usr/txtS_SPMON-LOW").text = zek1_range_left
        sess1.findById("wnd[0]/usr/txtS_SPMON-HIGH").text = zek1_range_right

        sess1.findById("wnd[0]").sendVKey(8)  # Load variant
        zek1_change_layout(session=sess1, layout_row_num="0")
        export_data_to_file_ZEK1(sess1, files_path, "cena_ZEK1.XLSX")

        show_message(f"Pliki zostały zapisane w katalogu {folder_name}.")

    except Exception as e:
        print("Błąd: ", e)
        error_details = traceback.format_exc()
        print("Szczegóły błędu:\n", error_details)
        input("Press Enter...")
