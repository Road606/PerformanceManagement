Attribute VB_Name = "test_ctrl_dashboard1"
Option Explicit


'Public obj_mdl_process_version As MDLProcessVersionMaster

Public Function setup()
    Dim obj_mdl_bin_prod_line As MDLBINProdLine

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

    hndl_log.init
    hndl_log.str_path = "C:\Users\czDanKle\Desktop\KLD\under-construction\app\performance\log\"
    hndl_log.str_file_name = "log.xlsx"
    hndl_log.open_data
End Function

Public Function tear_down()
    hndl_log.close_data
    
    Application.DisplayAlerts = True
End Function

Public Function test_run()
    Dim dbl_start As Double
    Dim dbl_end As Double
    
    Dim obj_dashboard As CtrlDashboard1
    
    setup
    
    dbl_start = Now
    Set obj_dashboard = New CtrlDashboard1
    ' run setting
    obj_dashboard.bool_run_process = new_const_ctrl_dashboard1.BOOL_RUN_PROCESS_YES 'new_const_ctrl_dashboard1.BOOL_RUN_PROCESS_YES
    'obj_dashboard.bool_run_process_load_unfinished = new_const_ctrl_dashboard1.BOOL_RUN_PROCESS_LOAD_UNFINISHED_NO
    obj_dashboard.bool_run_kpi_pallet = new_const_ctrl_dashboard1.BOOL_RUN_KPI_PALLET_YES
    
    obj_dashboard.before_run
    obj_dashboard.run
    dbl_end = Now
    
    tear_down
    
    Debug.Print Format(dbl_end - dbl_start, "HH:MM:SS")
End Function





