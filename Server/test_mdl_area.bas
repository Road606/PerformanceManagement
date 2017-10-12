Attribute VB_Name = "test_mdl_area"
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

Public Function test_area_data()
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    Dim obj_mdl_area As MDLAreaMD
    
    setup
    
    dbl_start = Now
    Set obj_mdl_area = New MDLAreaMD
    obj_mdl_area.init
    obj_mdl_area.STR_WS_NAME = "db.md.area"
    obj_mdl_area.load_data
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function

Public Function test_area_step_data()
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    Dim obj_mdl_area As MDLAreaMD
    Dim obj_mdl_area_step As MDLAreaStepMD
    
    setup
    
    dbl_start = Now
    Set obj_mdl_area = New MDLAreaMD
    obj_mdl_area.init
    obj_mdl_area.STR_WS_NAME = "db.md.area"
    obj_mdl_area.load_data
    
    Set obj_mdl_area_step = New MDLAreaStepMD
    obj_mdl_area_step.init
    obj_mdl_area_step.STR_WS_NAME = "db.md.area.step"
    obj_mdl_area_step.load_data obj_mdl_area.col_areas
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function

