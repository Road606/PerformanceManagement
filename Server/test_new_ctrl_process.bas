Attribute VB_Name = "test_new_ctrl_process"
'Option Explicit
'
'Public Function setup()
'    bin.init
'    wc.init
'
'    new_mdl_data_process.init
'    new_mdl_data_process.init_model new_mdl_data_process.STR_DB_TYPE_FILE_EXCEL
'      ' finished
'    new_mdl_data_process.obj_model.obj_finished.BYTE_STEP_STATUS = new_db_process_step.BYTE_CLOSED
'    new_mdl_data_process.obj_model.obj_finished.str_ws_name = "data.process"
'      ' unfinished
'    new_mdl_data_process.obj_model.obj_unfinished.BYTE_STEP_STATUS = new_db_process_step.BYTE_OPEN
'    new_mdl_data_process.obj_model.obj_unfinished.str_ws_name = "data.process.unfinished"
'
'    new_ctrl_process_master.init
'    new_ctrl_process_master.str_ws_name = "data.master.process"
'
'    new_ctrl_process_master_version.init
'    new_ctrl_process_master_version.str_ws_name = "data.master.process.version"
'
'    new_ctrl_pallet.init
'
'    new_mdl_history.init
'    new_mdl_history.STR_PATH_INBOUND = ThisWorkbook.Path & "\data\inbound\"
'    new_mdl_history.str_ws_name = "Sheet1"
'    new_mdl_history.str_file_appendix = ".xls"
'
'    new_file_processed_level1.STR_PATH_INBOUND = ThisWorkbook.Path & "\log\"
'    new_file_processed_level1.str_file_name = "file_processed.xlsx"
'
'    new_file_processed_level1.open_data
'
'    hndl_log.init
'    hndl_log.str_path = "C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\log\"
'    hndl_log.str_file_name = "log.xlsx"
'    hndl_log.open_data
'
'    Application.DisplayAlerts = False
'End Function
'
'Public Function tear_down()
'    new_file_processed_level1.close_data
'    hndl_log.close_data
'
'    Application.DisplayAlerts = True
'End Function
'
'Public Function test_load_multi()
'    Dim dbl_start As Double
'    Dim dbl_end As Double
'
'    setup
'
'    dbl_start = Now
'    new_mdl_data_process.obj_model.open_connection
'    ' load process data
'    new_ctrl_process_master.load_data
'
'    new_mdl_data_process.obj_model.obj_unfinished.load_multi
'
'    new_mdl_data_process.obj_model.close_connection
'    dbl_end = Now
'    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
'
'    tear_down
'End Function
'
'Public Function test_object()
'    Dim obj_test As CtrlHistoryToProcess
'
'    Set obj_test = New CtrlHistoryToProcess
'
'End Function
