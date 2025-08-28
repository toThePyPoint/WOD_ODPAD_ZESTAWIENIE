import os
import time
from other_functions import close_excel_file


def simple_load_variant(obj_sess, variant_name, open_only=False):
    obj_sess.findById("wnd[0]").sendVKey(17)
    obj_sess.findById("wnd[1]/usr/txtV-LOW").text = variant_name
    obj_sess.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    obj_sess.findById("wnd[1]/usr/txtV-LOW").caretPosition = 9
    obj_sess.findById("wnd[1]/tbar[0]/btn[8]").press()

    if open_only:
        return

    obj_sess.findById("wnd[0]").sendVKey(8)


def open_one_transaction(session, transaction_name):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n" + transaction_name
    session.findById("wnd[0]").sendVKey(0)


def paste_production_orders_and_load_variant(session):
    """
    Works for COHV and COOIS transactions.
    :param session:
    :return:
    """
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()


def export_data_to_file_MB51(session, file_path, file_name):
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = file_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(2)

    for _ in range(50):
        if close_excel_file(file_name):
            break
        time.sleep(0.1)


def export_data_to_file_ZEK1(session, file_path, file_name):
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = file_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(2)

    for _ in range(50):
        if close_excel_file(file_name):
            break
        time.sleep(0.1)


def export_data_to_file_COHV(session, file_path, file_name):

    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = file_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(2)

    for _ in range(50):
        if close_excel_file(file_name):
            break
        time.sleep(0.1)


def clear_sap_warnings(session):
    """
    Check for SAP warning messages and clear them if present.
    """
    try:
        message_bar = session.findById("wnd[0]/sbar")
        if message_bar.MessageType == "W":  # 'W' stands for Warning (Yellow message)
            session.findById("wnd[0]").sendVKey(0)  # Press Enter to acknowledge the warning
            time.sleep(0.2)  # Give SAP some time to process
            # print("SAP warning cleared.")
    except Exception as e:
        print(f"Error handling SAP message: {e}")


def get_sap_message(session):
    """
    Retrieve the text of the SAP message from the status bar.
    """
    try:
        message_bar = session.findById("wnd[0]/sbar")
        return message_bar.Text  # Return the message text
    except Exception as e:
        print(f"Error retrieving SAP message: {e}")
        return None  # Return None if there's an error


def zek1_change_layout(session, layout_row_num):
    session.findById("wnd[0]/tbar[1]/btn[33]").press()
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = str(layout_row_num)
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()