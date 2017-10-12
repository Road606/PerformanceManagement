Attribute VB_Name = "test_mdl_kpi"
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

Public Function test_kpi()
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    Dim obj_mdl_kpi As MDLKPIMD
    
    setup
    
    dbl_start = Now
    Set obj_mdl_process = New MDLProcessMaster
    obj_mdl_process.init
    obj_mdl_process.STR_WS_NAME = "data.master.process"
    Set obj_mdl_process_version = New MDLProcessVersionMaster
'    obj_mdl_process_version.init
    obj_mdl_process_version.STR_WS_NAME = "data.master.process.version"
    obj_mdl_process.load_data
    obj_mdl_process_version.load_data obj_mdl_process.col_process_masters
    
    
    Set obj_mdl_kpi = New MDLKPIMD
    obj_mdl_kpi.init
    obj_mdl_kpi.STR_WS_NAME = "db.md.kpi"
    obj_mdl_kpi.load_data obj_mdl_process.col_process_masters
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function

Public Function test_kpi_limit()
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    Dim obj_mdl_kpi As MDLKPIMD
    Dim obj_mdl_kpi_limit As MDLKPIOnTimeLimitsMD
    
    setup
    
    dbl_start = Now
    Set obj_mdl_process = New MDLProcessMaster
    obj_mdl_process.init
    obj_mdl_process.STR_WS_NAME = "data.master.process"
    Set obj_mdl_process_version = New MDLProcessVersionMaster
'    obj_mdl_process_version.init
    obj_mdl_process_version.STR_WS_NAME = "data.master.process.version"
    obj_mdl_process.load_data
    obj_mdl_process_version.load_data obj_mdl_process.col_process_masters
    
    
    Set obj_mdl_kpi = New MDLKPIMD
    obj_mdl_kpi.init
    obj_mdl_kpi.STR_WS_NAME = "db.md.kpi"
    obj_mdl_kpi.load_data obj_mdl_process.col_process_masters
    
    Set obj_mdl_kpi_limit = New MDLKPIOnTimeLimitsMD
'    obj_mdl_kpi_limit.init
    obj_mdl_kpi_limit.STR_WS_NAME = "db.md.kpi.on.time.limits"
    obj_mdl_kpi_limit.load_data obj_mdl_kpi.col_kpis
    
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function
