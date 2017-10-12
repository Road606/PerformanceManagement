Attribute VB_Name = "test_new_mdl_history"
Option Explicit

Public obj_mdl_data_process_finished As MDLDataProcessExcel
Public obj_mdl_data_process_unfinished As MDLDataProcessExcel
Public obj_mdl_bin_prod_line As MDLBINProdLine

Public Function setup()
    Dim obj_data_provider_util_mdl_history As FileExcelDataProviderUtil

    Dim obj_multi_data_provider As MultiFileExcelDataProvider
    Dim obj_data_provider As FileExcelDataProvider
    Dim obj_data_provider_util As FileExcelDataProviderUtil

    ' app
    hndl_log.init
    hndl_log.str_path = ThisWorkbook.Path & "\log\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\log\"
    hndl_log.str_file_name = "log.xlsx"
    hndl_log.open_data

    bin.init
    Set obj_mdl_bin_prod_line = New MDLBINProdLine
    Set obj_mdl_bin_prod_line.obj_data_provider = New FileExcelDataProvider
    obj_mdl_bin_prod_line.obj_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    obj_mdl_bin_prod_line.obj_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\master\"
    obj_mdl_bin_prod_line.obj_data_provider.str_file_name = "bin_prod_line.xlsx"
    obj_mdl_bin_prod_line.obj_data_provider.STR_WS_NAME = "data"
    obj_mdl_bin_prod_line.bool_load_into_local_collection = True
    obj_mdl_bin_prod_line.load
    Set bin.col_bin_prod_lines = obj_mdl_bin_prod_line.col_bin_prod_lines
    wc.init
    
    ' pallet controller
    new_ctrl_pallet.init
    ' data process model
      ' finished
    Set obj_mdl_data_process_finished = New MDLDataProcessExcel
    obj_mdl_data_process_finished.BYTE_STEP_STATUS = new_db_process_step.BYTE_CLOSED ' # implement byte_step_status should be evaluated in ctrl_pallet logic
    obj_mdl_data_process_finished.str_provider_id_default = "default"
        ' multi data provider
    Set obj_multi_data_provider = New MultiFileExcelDataProvider
    obj_multi_data_provider.STR_WS_NAME = "data" '"data.process"
    obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    obj_multi_data_provider.BOOL_IS_DYNAMIC_LOADING_PROVIDERS_FROM_FILES_ON = False
          ' file
    obj_multi_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-process\finished\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\"
    Set obj_mdl_data_process_finished.obj_multi_data_provider = obj_multi_data_provider
          ' processed file
    obj_multi_data_provider.STR_PROCESSED_FILE_PATH = ThisWorkbook.Path & "\log\"
    obj_multi_data_provider.STR_PROCESSED_FILE_NAME = "history_process_finished-file_processed.xlsx"
          ' add to pallet's controller
    Set new_ctrl_pallet.obj_mdl_finished = obj_mdl_data_process_finished
        ' data provider util
    Set obj_data_provider_util = New FileExcelDataProviderUtil
          ' save mode
    obj_data_provider_util.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY
          ' file
            ' tmpl
    obj_data_provider_util.str_tmpl_path = ThisWorkbook.Path & "\data\inbound\history-process\tmpl\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\tmpl\"
    obj_data_provider_util.str_tmpl_file_name_prefix = "history-process"
    obj_data_provider_util.str_tmpl_file_name_appendix = "tmpl-v0_2"
            ' new
    obj_data_provider_util.str_file_prefix = "history-process"
    obj_data_provider_util.str_file_name_separator = "-"
    obj_data_provider_util.str_file_appendix = ".xlsx"
    'Debug.Print obj_data_provider_util.retrieve_provider_id_reverse("history-process-17052203.xlsx")
    'Dim obj_period As Period
    'Set obj_period = obj_data_provider_util.retrieve_period("17052203")
          ' add to multi data provider
'    Set obj_mdl_data_process_finished.obj_data_provider_util = obj_data_provider_util
    Set obj_multi_data_provider.obj_data_provider_util = obj_data_provider_util
    
      ' unfinished
    Set obj_mdl_data_process_unfinished = New MDLDataProcessExcel
    obj_mdl_data_process_unfinished.BYTE_STEP_STATUS = new_db_process_step.BYTE_OPEN ' # implement byte_step_status should be evaluated in ctrl_pallet logic
    obj_mdl_data_process_unfinished.str_provider_id_default = "default"
    obj_mdl_data_process_unfinished.add_listener New ListenerProcessToPallet
        ' multi data provider
    Set obj_multi_data_provider = New MultiFileExcelDataProvider
    obj_multi_data_provider.STR_WS_NAME = "data" '"data.process"
    obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    obj_multi_data_provider.BOOL_IS_DYNAMIC_LOADING_PROVIDERS_FROM_FILES_ON = False
          ' file
    obj_multi_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-process\unfinished\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\"
            ' specific provider
    Set obj_data_provider = New FileExcelDataProvider
    obj_data_provider.str_id = "unfinished"
    obj_data_provider.STR_WS_NAME = "data"
    obj_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL
    obj_data_provider.str_path = ThisWorkbook.Path & "\data\inbound\history-process\unfinished\"
    obj_data_provider.str_file_name = "history-process-unfinished.xlsx"
    obj_data_provider.bool_save_mode_on = True
    obj_data_provider.add_listener obj_mdl_data_process_unfinished
    obj_multi_data_provider.add_provider obj_data_provider
    
    Set obj_mdl_data_process_unfinished.obj_multi_data_provider = obj_multi_data_provider
    obj_mdl_data_process_unfinished.str_static_data_provider_id = obj_data_provider.str_id
          ' processed file
    obj_multi_data_provider.STR_PROCESSED_FILE_PATH = ThisWorkbook.Path & "\log\"
    obj_multi_data_provider.STR_PROCESSED_FILE_NAME = "history_process_unfinished-file_processed.xlsx"
          ' add to pallet's controller
    Set new_ctrl_pallet.obj_mdl_unfinished = obj_mdl_data_process_unfinished
        ' data provider util
    Set obj_data_provider_util = New FileExcelDataProviderUtil
          ' save mode
    obj_data_provider_util.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY
          ' file
            ' tmpl
    obj_data_provider_util.str_tmpl_path = ThisWorkbook.Path & "\data\inbound\history-process\tmpl\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\tmpl\"
    obj_data_provider_util.str_tmpl_file_name_prefix = "history-process"
    obj_data_provider_util.str_tmpl_file_name_appendix = "tmpl-v0_2"
            ' new
    obj_data_provider_util.str_file_prefix = "history-process-unfinished"
    obj_data_provider_util.str_file_name_separator = "-"
    obj_data_provider_util.str_file_appendix = ".xlsx"
    'Debug.Print obj_data_provider_util.retrieve_provider_id_reverse("history-process-17052203.xlsx")
    'Dim obj_period As Period
    'Set obj_period = obj_data_provider_util.retrieve_period("17052203")
          ' add to multi data provider
'    Set obj_mdl_data_process_unfinished.obj_data_provider_util = obj_data_provider_util
    Set obj_multi_data_provider.obj_data_provider_util = obj_data_provider_util
    
'      ' unfinished
'    Set obj_mdl_data_process_unfinished = New MDLDataProcessExcel
'    obj_mdl_data_process_unfinished.BYTE_STEP_STATUS = new_db_process_step.BYTE_OPEN
'    'obj_mdl_data_process_unfinished.str_save_mode = obj_mdl_data_process_finished.STR_SAVE_MODE_HOURLY
'    obj_mdl_data_process_unfinished.add_listener New ListenerProcessToPallet
'        ' multi data provider
'    Set obj_multi_data_provider = New MultiFileExcelDataProvider
'          ' general
'    obj_multi_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_INTERNAL
'          ' sheet
'    obj_multi_data_provider.STR_WS_NAME = "data.process.unfinished"
'          ' processed file
'    obj_multi_data_provider.STR_PROCESSED_FILE_PATH = ThisWorkbook.Path & "\log\"
'    obj_multi_data_provider.STR_PROCESSED_FILE_NAME = "history_process_unfinished-file_processed.xlsx" '"file_processed.xlsx"
'          ' add to multi data provider
'    Set obj_mdl_data_process_unfinished.obj_multi_data_provider = obj_multi_data_provider
'        ' specific provider
'    Set obj_data_provider = New FileExcelDataProvider
'    obj_data_provider.STR_WS_NAME = "data.process.unfinished"
'    obj_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_INTERNAL
'    obj_data_provider.add_listener obj_mdl_data_process_unfinished
'    obj_mdl_data_process_unfinished.obj_multi_data_provider.add_provider obj_data_provider
'        ' data provider util
'    Set obj_data_provider_util = New FileExcelDataProviderUtil
'          ' save mode
'    obj_data_provider_util.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY
'          ' add to model
''    Set obj_mdl_data_process_unfinished.obj_data_provider_util = obj_data_provider_util
'    Set obj_multi_data_provider.obj_data_provider_util = obj_data_provider_util
'         ' add to pallet controller
'    Set new_ctrl_pallet.obj_mdl_unfinished = obj_mdl_data_process_unfinished
      
    ' master data for process
    new_ctrl_process_master.init
    new_ctrl_process_master.STR_WS_NAME = "data.master.process"
    new_ctrl_process_master_action.init
    new_ctrl_process_master_action.STR_WS_NAME = "data.master.process.action"
    new_ctrl_process_master_transac.init
    new_ctrl_process_master_transac.STR_WS_NAME = "data.master.process.transaction"
    new_ctrl_process_master_version.init
    new_ctrl_process_master_version.STR_WS_NAME = "data.master.process.version"
    new_ctrl_process_master_step.init
    new_ctrl_process_master_step.STR_WS_NAME = "data.master.process.ver.step"
    ' load process master data
    new_ctrl_process_master.load_data
    
    
    
    ' mdl history
    new_mdl_history.init
    new_mdl_history.STR_PATH_INBOUND = ThisWorkbook.Path & "\data\inbound\history-pallet\"
    new_mdl_history.STR_WS_NAME = "data" '"Sheet1"
    new_mdl_history.str_file_appendix = ".xls"
    new_mdl_history.STR_HBW_USER = "DK2WEBHBW"
    new_mdl_history.col_listeners.add New ListenerHistoryToProcess
      ' input
    Set obj_data_provider_util_mdl_history = New FileExcelDataProviderUtil
          ' save mode
    obj_data_provider_util_mdl_history.str_save_mode = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY
          ' file
'            ' tmpl
'    obj_data_provider_util.str_tmpl_path = ThisWorkbook.Path & "\data\inbound\history-process\tmpl\" '"C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\data\inbound\history-process\tmpl\"
'    obj_data_provider_util.str_tmpl_file_name_prefix = "history-process"
'    obj_data_provider_util.str_tmpl_file_name_appendix = "tmpl-v0_2"
            ' new
    obj_data_provider_util_mdl_history.str_file_prefix = "history_pallet"
    obj_data_provider_util_mdl_history.str_file_name_separator = "-"
    obj_data_provider_util_mdl_history.str_file_appendix = ".xlsx"
    
    Set new_mdl_history.obj_file_data_provider_util = obj_data_provider_util_mdl_history
    
    
    new_file_processed_level1.STR_PATH_INBOUND = ThisWorkbook.Path & "\log\"
    new_file_processed_level1.str_file_name = "history-file_processed.xlsx"
    new_file_processed_level1.open_data
    
    Application.DisplayAlerts = False
End Function

Public Function tear_down()
    new_file_processed_level1.close_data
    hndl_log.close_data
    
    Application.DisplayAlerts = True
End Function

Public Function test_process()
    Dim dbl_start As Double
    Dim dbl_end As Double

    setup
    
    dbl_start = Now
    
    obj_mdl_data_process_unfinished.set_clear_data
    obj_mdl_data_process_unfinished.load_static
    'obj_mdl_data_process_unfinished.obj_multi_data_provider.close_providers
'    new_mdl_data_process.obj_model.obj_unfinished.load_multi
    new_mdl_history.process
    ' make post process actions in listeners
      ' save unfinished
    obj_mdl_data_process_unfinished.reset_clear_data
    new_ctrl_pallet.save_open_pallets
    obj_mdl_data_process_finished.obj_multi_data_provider.close_providers
    obj_mdl_data_process_unfinished.obj_multi_data_provider.close_providers
    dbl_end = Now
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
    
    tear_down
End Function


