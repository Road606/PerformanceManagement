Attribute VB_Name = "test_file_provider"
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

Public Function test_load()
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    Dim obj_multi_provider As MultiFileExcelDataProvider
    
    setup
    
    dbl_start = Now
    Set obj_multi_provider = New MultiFileExcelDataProvider
    obj_multi_provider.str_path = "C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-pallet\files\"
    obj_multi_provider.str_file_appendix = ".xlsx"
    obj_multi_provider.STR_PROCESSED_FILE_PATH = "C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\log\"
    obj_multi_provider.STR_PROCESSED_FILE_NAME = "file_processed.xlsx"
    obj_multi_provider.STR_WS_NAME = "data"
    obj_multi_provider.BOOL_IS_DYNAMIC_LOADING_PROVIDERS_FROM_FILES_ON = True
    obj_multi_provider.load_data
    
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function

