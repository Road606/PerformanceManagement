Attribute VB_Name = "test_file_provider_util"
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

Public Function test_retrieve_period()
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    Dim obj_data_provider_util As FileExcelDataProviderUtil
    
    setup
    
    dbl_start = Now
    Set obj_data_provider_util = New FileExcelDataProviderUtil
          ' save mode
    obj_data_provider_util.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY
          ' file
            ' tmpl
    obj_data_provider_util.str_tmpl_path = "C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\tmpl\"
    obj_data_provider_util.str_tmpl_file_name_appendix = "tmpl-v0_1"
            ' new
    obj_data_provider_util.str_file_prefix = "history-process"
    obj_data_provider_util.str_file_name_separator = "-"
    obj_data_provider_util.str_file_appendix = ".xlsx"
    Debug.Print obj_data_provider_util.retrieve_provider_id_reverse("history-process-17052203.xlsx")
    Dim obj_period As Period
    Set obj_period = obj_data_provider_util.retrieve_period("17052203")
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function
