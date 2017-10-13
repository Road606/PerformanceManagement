Attribute VB_Name = "test_MD_Dashboard1"
Option Explicit

Public Function setup()
    hndl_log.init
    hndl_log.str_path = "C:\Users\czjirost\Desktop\"
    hndl_log.str_file_name = "log.xlsx"
    hndl_log.open_data
End Function

Public Function tear_down()
    hndl_log.close_data
    
    Application.DisplayAlerts = True
End Function

Public Function test_md_dashboard_data()
    Dim obj_dashboard As New MDDashboard1
    
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    setup
    
    dbl_start = Now
    
    obj_dashboard.load
    
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
    
End Function
