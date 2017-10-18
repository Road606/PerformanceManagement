Attribute VB_Name = "new_mdl_history"
Option Explicit

Public Const str_module As String = "new_mdl_history"

Public STR_PATH_INBOUND As String
Public STR_PATH_OUTBOUND As String
Public str_file_appendix As String

' local data
Public STR_WS_NAME As String
Public STR_DATA_FIRST_CELL As String
Public Const STR_FIRST_ROW_TBL As String = "A1:AH1"
Public Const STR_FIRST_ROW_DATA As String = "A2:AH2"

' objects


' status bar
Public Const STR_STATUS_BAR_PREFIX As String = "Data history->"
'Public Const STR_STATUS_BAR_PREFIX_UPLOAD As String = "Uploading->"
Public Const STR_STATUS_BAR_PREFIX_PROCESS_DATA As String = "Processing data->"
Public Const LNG_STATUS_BAR_REFRESH As Long = 100

Public wb As Workbook

Public col_listeners As Collection

Public obj_file_data_provider_util As FileExcelDataProviderUtil

Public STR_HBW_USER As String

Public Function init()
    ' local
    STR_DATA_FIRST_CELL = "A2"
    ' external
    new_file_processed_level1.init
    
    Set col_listeners = New Collection
End Function

Public Function process()
    Dim obj_list_files As Collection
    Dim var_file As Variant
    Dim obj_data_provider_info As FileExcelDataProviderInfo
    Dim obj_listener As Object
    
    On Error GoTo ERR_RETRIEVE_FILES
    Set obj_list_files = retrieve_files
    On Error GoTo 0
    
    For Each var_file In obj_list_files
        Debug.Print STR_PATH_INBOUND & var_file
        Set obj_data_provider_info = New FileExcelDataProviderInfo
        obj_data_provider_info.str_provider_id = obj_file_data_provider_util.retrieve_provider_id_reverse(CStr(var_file))
        Set obj_data_provider_info.obj_period = obj_file_data_provider_util.retrieve_period(obj_data_provider_info.str_provider_id)
        For Each obj_listener In col_listeners
            obj_listener.before_process_history_record obj_data_provider_info
        Next
        
        'On Error GoTo ERR_PROCESS_FILE
        process_file CStr(var_file)
        On Error GoTo 0
    Next
            
    Exit Function
ERR_RETRIEVE_FILES:
    Debug.Print Err.Number & "->" & Err.description
    hndl_log.log db_log.TYPE_WARN, str_module, "process", Err.Number & "->" & Err.description
    Exit Function
ERR_PROCESS_FILE:
    hndl_log.log db_log.TYPE_WARN, str_module, "process", _
        "An error occured during processing file " & var_file & ". Original error message>" & Err.description
    Resume Next
End Function

Public Function process_file(STR_FILE As String)
    Dim rg_record As Range
    Dim obj_history_record As DBHistoryRecord
    Dim obj_listener As Object
    
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_PROCESS_DATA & "Opening history of file: " & STR_PATH_INBOUND & STR_FILE & "."
    
    'Set wb = util_file.open_wb(STR_PATH_INBOUND & str_file)
    Set wb = open_file(STR_FILE)
    'Set ws_local = wb.Worksheets(STR_WS_NAME)
    
    app_status_bar.str_module = STR_STATUS_BAR_PREFIX
    app_status_bar.str_method = STR_STATUS_BAR_PREFIX_PROCESS_DATA
    app_status_bar.STR_FILE = STR_FILE
    app_status_bar.lng_refresh_num = LNG_STATUS_BAR_REFRESH
    
    ' process records
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    app_status_bar.lng_total_records = rg_record.End(xlDown).Row
    
    Do While rg_record.Value <> ""
        DoEvents
        
        Set obj_history_record = create_history_record(rg_record)
        For Each obj_listener In col_listeners
            obj_listener.process_history_record obj_history_record
        Next
        
        app_status_bar.update_records rg_record.Row
        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop
        
    
    
    ' close data and make record this file was processed
    close_file wb, STR_FILE
End Function

Public Function create_history_record(rg_record As Range) As DBHistoryRecord
    Set create_history_record = New DBHistoryRecord
    
    create_history_record.str_transaction_started = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
    create_history_record.str_transaction_finished = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
    
    create_history_record.str_transaction_type_started = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TYPE_STARTED).Value
    create_history_record.str_transaction_type_finished = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TYPE_FINISHED).Value
    
    create_history_record.str_transaction_code = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_CODE).Value
    
    create_history_record.str_combi_vhu_from = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_FROM).Value
    create_history_record.str_combi_vhu_to = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value
    
    create_history_record.str_bin_from = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_FROM).Value
    create_history_record.str_bin_to = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
    
    create_history_record.str_material = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_MATERIAL).Value
    create_history_record.str_gr_vendor = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_GR_VENDOR).Value
    
    create_history_record.lng_qty_from = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_QTY_FROM).Value
    create_history_record.lng_qty_to = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_QTY_TO).Value
    create_history_record.dbl_pallet_utilization = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_PALLET_UTILIZATION).Value
    
    create_history_record.str_hu_status_from = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_HU_STATUS_FROM).Value
    create_history_record.str_hu_status_to = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_HU_STATUS_TO).Value
    
    create_history_record.str_stock_type_from = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_STOCK_TYPE_FROM).Value
    create_history_record.str_stock_type_to = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_STOCK_TYPE_TO).Value
    
    create_history_record.str_gr_type = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_GR_TYPE).Value
    create_history_record.str_gr_vendor = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_GR_VENDOR).Value
    create_history_record.str_gr_order = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_GR_ORDER).Value
    create_history_record.str_gr_datetime = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_GR_DATETIME).Value
    
    create_history_record.str_storage_group_material = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_STORAGE_GROUP_MATERIAL).Value
    create_history_record.str_order_delivery = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_ORDER_DELIVERY).Value
    create_history_record.str_wc_ship_to = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_WC_SHIPTO).Value
    create_history_record.str_machine_transport_ref = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_MACHINE).Value
    
    create_history_record.str_user_name = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_USERNAME).Value
    
    create_history_record.str_task_list_type = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TASK_LIST_TYPE).Value
    
    create_history_record.str_shipping_type = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_SHIPPING_TYPE).Value

'      ' hu split
'    rg_row.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_HU_SOURCE).NumberFormat = "0"
'    rg_row.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_HU_SOURCE).Value = obj_record.str_hu_source
'    rg_row.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_HU_DESTINATION).NumberFormat = "0"
'    rg_row.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_HU_DESTINATION).Value = obj_record.str_hu_destination

'
'    ' transaction data
'    rg_row.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_DATE).Value = Format(DateValue(obj_record.str_transaction_started), db_history_record.STR_TRANSACTION_DATE_FORMAT)
'    rg_row.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TIME).NumberFormat = "hh:mm:ss" 'db_history_record.STR_TRANSACTION_TIME_FORMAT
'    rg_row.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TIME).Value = Format(TimeValue(obj_record.str_transaction_started), db_history_record.STR_TRANSACTION_TIME_FORMAT)
    
End Function

'Public Function create_history_record(rg_record As Range) As DBHistoryRecord
'    Set create_history_record = New DBHistoryRecord
'
'    create_history_record.str_transaction_started = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'    create_history_record.str_transaction_finished = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
'
'    create_history_record.str_transaction_type_started = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TYPE_STARTED).Value
'    create_history_record.str_transaction_type_finished = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TYPE_FINISHED).Value
'
'    create_history_record.str_transaction_code = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_CODE).Value
'
'    create_history_record.str_combi_vhu_from = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_FROM).Value
'    create_history_record.str_combi_vhu_to = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value
'
'    create_history_record.str_bin_from = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_FROM).Value
'    create_history_record.str_bin_to = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
'
'    'create_history_record.str_bin_type_from = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_FROM_TYPE).Value
'    'create_history_record.str_bin_type_to = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO_TYPE).Value
'
'    create_history_record.str_material = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_MATERIAL).Value
'
'    create_history_record.str_wc_shipto = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_WC_SHIPTO).Value
'
'    create_history_record.str_user_name = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_USERNAME).Value
'
'    create_history_record.str_task_list_type = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TASK_LIST_TYPE).Value
'End Function

Public Function open_file(str_file_name As String) As Workbook
    Set open_file = util_file.open_wb(STR_PATH_INBOUND & str_file_name)
    ' create zg804 file record - serves as evidence that file was already processed
    Set new_file_processed_level1.obj_file_record = New DBProcessedFileRecord
    new_file_processed_level1.obj_file_record.str_name = str_file_name
    new_file_processed_level1.obj_file_record.str_date_started = Now
End Function

Public Function close_file(wb As Workbook, str_file_name As String)
    wb.Close SaveChanges:=False
    new_file_processed_level1.obj_file_record.str_date_finished = Now
    new_file_processed_level1.save_record
    
    On Error GoTo ERR_MOVE_FILE
    'Name (STR_PATH_INBOUND & str_file_name) As (STR_PATH_OUTBOUND & str_file_name)
    On Error GoTo 0
    Exit Function
ERR_MOVE_FILE:
    Debug.Print "hndl_history->close_file->move file from: " & (STR_PATH_INBOUND & str_file_name) & " to " & (STR_PATH_OUTBOUND & str_file_name)
End Function

Public Function retrieve_files() As Collection
    Dim rg_processed_file As Range
    
    Set retrieve_files = util_file.retrieve_files(STR_PATH_INBOUND, "*" & str_file_appendix)
    Set rg_processed_file = new_file_processed_level1.ws.Range(new_file_processed_level1.STR_DATA_START_RG)
    
    Do While rg_processed_file.Value <> ""
        On Error GoTo INFO_NOT_PROCESSED_FILE
        retrieve_files.Remove rg_processed_file
        On Error GoTo 0
        Set rg_processed_file = rg_processed_file.Offset(1)
    Loop
    Exit Function
INFO_NOT_PROCESSED_FILE:
    Resume Next
End Function



