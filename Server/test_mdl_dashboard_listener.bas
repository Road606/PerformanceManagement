Attribute VB_Name = "test_mdl_dashboard_listener"
Option Explicit

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

Public Function test_kpi()
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    Dim obj_md_dashboard As MDDashboard1
    Dim obj_file_excel_provider As FileExcelDataProvider
    Dim obj_mdl_listener As MDLProcessDashboardListener
    
    setup
    
    dbl_start = Now
    
    Set obj_md_dashboard = New MDDashboard1
    obj_md_dashboard.load
    
    Set obj_mdl_listener = New MDLProcessDashboardListener
    Set obj_mdl_listener.obj_md_dashboard = obj_md_dashboard
    obj_mdl_listener.load
    
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function


