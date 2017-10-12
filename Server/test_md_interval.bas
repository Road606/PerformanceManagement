Attribute VB_Name = "test_md_interval"
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

Public Function test_creation()
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    Dim obj_mdl_interval As MDLProcessIntervalMD
    Dim obj_data_provider As FileExcelDataProvider
    
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
    
    Set obj_mdl_interval = New MDLProcessIntervalMD
    Set obj_mdl_interval.col_processes = obj_mdl_process.col_process_masters
    Set obj_data_provider = New FileExcelDataProvider
    obj_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_INTERNAL
    obj_data_provider.STR_WS_NAME = "db.md.process.interval"
    obj_data_provider.add_listener obj_mdl_interval
    Set obj_mdl_interval.obj_data_provider = obj_data_provider
    obj_mdl_interval.load
    
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function
