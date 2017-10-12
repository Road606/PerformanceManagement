Attribute VB_Name = "test_md_dashboard"
Option Explicit


'Public obj_mdl_process_version As MDLProcessVersionMaster

Public Function setup()
    hndl_log.init
    hndl_log.str_path = "C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\log\"
    hndl_log.str_file_name = "log.xlsx"
    hndl_log.open_data
End Function

Public Function tear_down()
    hndl_log.close_data
    
    Application.DisplayAlerts = True
End Function

Public Function test_load()
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    Dim obj_md_dashboard As MDDashboard1
    
    setup
    
    dbl_start = Now
    Set obj_md_dashboard = New MDDashboard1
    obj_md_dashboard.load
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function




