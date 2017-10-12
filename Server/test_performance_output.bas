Attribute VB_Name = "test_performance_output"
Option Explicit

Public Function setup()
    app.init
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
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = ""
    
    hndl_log.open_data
    hndl_history_file_processed.open_data
    
    'hndl_proc_in_ra_vna_rack.STR_DURATION_UNIT = "n" ' n = minutes, s = seconds
End Function

Public Function tear_down()
    
    hndl_log.close_data
    hndl_history_file_processed.close_data
        
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
End Function

Public Function test_save_output()
    Dim dbl_start As Double
    Dim dbl_end As Double

    'app.before_run
    setup
    
    dbl_start = Now
    hndl_performance_output.save
    dbl_end = Now
    
    'app.after_run
    tear_down
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function

