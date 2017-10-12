Attribute VB_Name = "test_mdl_process_master"
Option Explicit

Public obj_mdl_process As MDLProcessMaster
Public obj_mdl_process_version As MDLProcessVersionMaster

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
    
'    new_ctrl_process_master.init
'    new_ctrl_process_master.STR_WS_NAME = "data.master.process"
'
'    new_ctrl_process_master_action.init
'    new_ctrl_process_master_action.STR_WS_NAME = "data.master.process.action"
'
'    new_ctrl_process_master_transac.init
'    new_ctrl_process_master_transac.STR_WS_NAME = "data.master.process.transaction"
'
'    new_ctrl_process_master_version.init
'    new_ctrl_process_master_version.STR_WS_NAME = "data.master.process.version"
'
'    new_ctrl_process_master_step.init
'    new_ctrl_process_master_step.STR_WS_NAME = "data.master.process.step"
    Set obj_mdl_process = New MDLProcessMaster
    obj_mdl_process.init
    obj_mdl_process.STR_WS_NAME = "data.master.process"
    
    Set obj_mdl_process_version = New MDLProcessVersionMaster
'    obj_mdl_process_version.init
    obj_mdl_process_version.STR_WS_NAME = "data.master.process.version"
    
    hndl_log.init
    hndl_log.str_path = "C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\log\"
    hndl_log.str_file_name = "log.xlsx"
    hndl_log.open_data
End Function

Public Function tear_down()
    hndl_log.close_data
    
    Application.DisplayAlerts = True
End Function

Public Function test_process_data()
    Dim dbl_start As Double
    Dim dbl_end As Double

    'app.before_run
    setup
    
    dbl_start = Now
    obj_mdl_process.load_data
    dbl_end = Now
    
    tear_down
    'app.after_run
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function

Public Function test_process_version_data()
    Dim dbl_start As Double
    Dim dbl_end As Double

    'app.before_run
    setup
    
    dbl_start = Now
    obj_mdl_process.load_data
    obj_mdl_process_version.load_data obj_mdl_process.col_process_masters
    dbl_end = Now
    
    tear_down
    'app.after_run
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function

