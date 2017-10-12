Attribute VB_Name = "test_mdl_place"
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

Public Function test_place_data()
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    Dim obj_mdl_place As MDLPlaceMD
    
    setup
    
    dbl_start = Now
    Set obj_mdl_place = New MDLPlaceMD
    obj_mdl_place.init
    obj_mdl_place.STR_WS_NAME = "db.md.place"
    obj_mdl_place.load_data
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function



