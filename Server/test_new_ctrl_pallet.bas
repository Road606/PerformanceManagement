Attribute VB_Name = "test_new_ctrl_pallet"
Option Explicit

Public Function setup()
    'app.init
    'bin.init
    'hndl_history.init
    'hndl_performance.init
    'hndl_performance_output.init
    'hndl_proc_in_ra_vna_rack.init
    'hndl_proc_inbound_vna_in_rack.init
    
    'hndl_proc_in_ra_vna_rack.init
    'hndl_proc_inbound_vna_in_rack.init
    
    'hndl_history.STR_PATH_INBOUND = ThisWorkbook.Path & "\data\inbound\"
    'hndl_history.STR_PATH_OUTBOUND = ThisWorkbook.Path & "\data\inbound\processed\"
    
    'hndl_performance_output.STR_OUTBOUND_PATH = ThisWorkbook.Path & "\data\outbound\"
    'hndl_performance_output.STR_OUTBOUND_TMPL_PATH = ThisWorkbook.Path & "\tmpl\"
    'hndl_performance_output.STR_PASSWD = "db_history"
    
    'hndl_proc_in_ra_vna_rack.STR_DURATION_UNIT = "n" ' n = minutes, s = seconds
    
    new_ctrl_pallet.init
    'new_ctrl_pallet.STR_WS_NAME = "data.pallet"
End Function

Public Function test_process_data()
    Dim dbl_start As Double
    Dim dbl_end As Double

    'app.before_run
    setup
    
    dbl_start = Now
    'new_ctrl_pallet.load
    dbl_end = Now
    
    'app.after_run
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function

