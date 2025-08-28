def cohv_go_back_change_layout_list_type_and_order_type(session, list_type, layout, ord_type):
    session.findById("wnd[0]/tbar[0]/btn[3]").press()  # Go back
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = list_type
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = layout
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_AUART-LOW").text = ord_type
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(8)
