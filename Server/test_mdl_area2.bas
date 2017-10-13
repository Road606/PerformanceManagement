Attribute VB_Name = "test_mdl_area2"
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

Public Function test_place_data()
    Dim dbl_start As Double
    Dim dbl_end As Double
    Dim test_collection As New Collection
    Dim listener As New DummyListener
    
    Dim MDLArea As New MDLAreaMD
    MDLArea.single_data_provider.STR_DATA_FIRST_CELL = "A2"
    MDLArea.single_data_provider.STR_WS_NAME = "db.md.area"
    
    setup
    
    dbl_start = Now
    
    MDLArea.add_listener listener
    MDLArea.load_data
    
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
    
    
End Function
