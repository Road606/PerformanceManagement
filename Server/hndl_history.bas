Attribute VB_Name = "hndl_history"
Option Explicit

Public Const str_module As String = "hndl_history"

Public STR_PATH_INBOUND As String
Public STR_PATH_OUTBOUND As String
Public str_file_appendix As String

' local data
Public STR_WS_NAME As String
Public STR_DATA_FIRST_CELL As String
Public Const STR_FIRST_ROW_TBL As String = "A1:AG1"
Public Const STR_FIRST_ROW_DATA As String = "A2:AG2"

' objects


' status bar
Public Const STR_STATUS_BAR_PREFIX As String = "Data history->"
'Public Const STR_STATUS_BAR_PREFIX_UPLOAD As String = "Uploading->"
Public Const STR_STATUS_BAR_PREFIX_PROCESS_DATA As String = "Processing data->"
Public Const LNG_STATUS_BAR_REFRESH As Long = 100

Public wb As Workbook

Public Function init()
    'STR_PATH_INBOUND = "C:\Users\czDanKle\Desktop\KLD\under-construction\wh-map\app\data\inbound\zg804\unprocessed\"
    'STR_PATH_OUTBOUND = "C:\Users\czDanKle\Desktop\KLD\under-construction\wh-map\app\data\inbound\zg804\processed\"
    str_file_appendix = ".xls"
    
    STR_WS_NAME = "data"
    STR_DATA_FIRST_CELL = "A2"
End Function

Public Function process_data()
    Dim obj_list_files As Collection
    Dim var_file As Variant
    
    On Error GoTo ERR_RETRIEVE_FILES
    'Set obj_list_files = util_file.retrieve_files(STR_PATH_INBOUND, "*" & STR_FILE_APPENDIX)
    Set obj_list_files = retrieve_files
    On Error GoTo 0
    
    For Each var_file In obj_list_files
        Debug.Print STR_PATH_INBOUND & var_file
        'On Error GoTo ERR_PROCESS_FILE
        process_file CStr(var_file)
        'On Error GoTo 0
    Next
    
    Debug.Print "Number of incomplete records > " & hndl_process.obj_records.Count
    Dim pi As Object
    
    For Each pi In hndl_process.obj_records
        Set hndl_log.obj_log_record = New DBLogRecord
        hndl_log.obj_log_record.str_datetime = Now
        hndl_log.obj_log_record.str_type = db_log.TYPE_INFO
        hndl_log.obj_log_record.str_module = str_module
        hndl_log.obj_log_record.str_function = "process_data->incomplete"
        hndl_log.obj_log_record.str_message = "Incomplete movements> " & _
            pi.get_start_datetime & ">" & pi.str_pallet_id & ">" & pi.str_process_type & ";" & pi.str_process_subtype & ";" & pi.str_process_step
        hndl_log.save_record
    Next
    
    hndl_performance_output.save
    hndl_performance.clear
    Exit Function
ERR_RETRIEVE_FILES:
    Debug.Print Err.Number & "->" & Err.description
    Exit Function
ERR_PROCESS_FILE:
    Set hndl_log.obj_log_record = New DBLogRecord
    hndl_log.obj_log_record.str_datetime = Now
    hndl_log.obj_log_record.str_type = db_log.TYPE_ERR
    hndl_log.obj_log_record.str_module = hndl_history.str_module
    hndl_log.obj_log_record.str_function = "process_data"
    hndl_log.obj_log_record.str_message = "An error occured during processing file " & var_file & ". Original error message>" & Err.description
    hndl_log.save_record
    Resume Next
End Function

Public Function process_file(STR_FILE As String)

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_PROCESS_DATA & "opening history of file: " & STR_PATH_INBOUND & STR_FILE & "."
    
    'Set wb = util_file.open_wb(STR_PATH_INBOUND & str_file)
    Set wb = open_data(STR_FILE)
    'Set ws_local = wb.Worksheets(STR_WS_NAME)
        
    app_status_bar.str_module = ""
    
    app_status_bar.str_module = STR_STATUS_BAR_PREFIX
    app_status_bar.str_method = STR_STATUS_BAR_PREFIX_PROCESS_DATA
    app_status_bar.STR_FILE = STR_FILE
    app_status_bar.lng_refresh_num = LNG_STATUS_BAR_REFRESH
    
    process_file_ascending wb
    'process_file_descending wb
        
    ' save output
    hndl_performance_output.save
    hndl_performance.clear
    
    ' close data and make record this file was processed
    close_data wb, STR_FILE
End Function

Public Function process_file_ascending(wb As Workbook)
    Dim rg_record As Range

    sort_ascending wb
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    app_status_bar.lng_total_records = rg_record.End(xlDown).Row
    
    Do While rg_record.Value <> ""
        DoEvents
                        
        'If hndl_history_record.is_relevant(rg_record) Then ' check if this report is interested in current transaction
        hndl_process.run rg_record
        'End If
                                
        app_status_bar.update_records rg_record.Row
        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop
End Function

Public Function process_file_descending(wb As Workbook)
    Dim rg_record As Range

    sort_descending wb
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    app_status_bar.lng_total_records = rg_record.End(xlDown).Row
    
    Do While rg_record.Value <> ""
        DoEvents
                        
        If hndl_history_record.is_relevant(rg_record) Then ' check if this report is interested in current transaction
            'hndl_proc_inbound_vna_in_rack.run rg_record
        End If
                                
        app_status_bar.update_records rg_record.Row
        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop
End Function

Public Function open_data(str_file_name As String) As Workbook
    Set open_data = util_file.open_wb(STR_PATH_INBOUND & str_file_name)
    ' create zg804 file record - serves as evidence that file was already processed
    Set hndl_history_file_processed.obj_file_record = New DBProcessedFileRecord
    hndl_history_file_processed.obj_file_record.str_name = str_file_name
    hndl_history_file_processed.obj_file_record.str_date_started = Now
End Function

Public Function close_data(wb As Workbook, str_file_name As String)
    wb.Close SaveChanges:=False
    hndl_history_file_processed.obj_file_record.str_date_finished = Now
    hndl_history_file_processed.save_record
    
    On Error GoTo ERR_MOVE_FILE
    'Name (STR_PATH_INBOUND & str_file_name) As (STR_PATH_OUTBOUND & str_file_name)
    On Error GoTo 0
    Exit Function
ERR_MOVE_FILE:
    Debug.Print "hndl_history->close_data->move file from: " & (STR_PATH_INBOUND & str_file_name) & " to " & (STR_PATH_OUTBOUND & str_file_name)
End Function

Public Function retrieve_files() As Collection
    Dim rg_processed_file As Range
    
    Set retrieve_files = util_file.retrieve_files(STR_PATH_INBOUND, "*" & str_file_appendix)
    Set rg_processed_file = hndl_history_file_processed.ws.Range(hndl_history_file_processed.STR_DATA_START_RG)
    
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

Public Function sort_ascending(wb As Workbook)
    Dim ws As Worksheet
    Dim rg As Range
    'Dim rg_sort As Range
    
    Set ws = wb.Worksheets(STR_WS_NAME)
    Set rg = ws.Range(STR_DATA_FIRST_CELL)
    
    'ws_local.Activate
    'ws.Range(rg.Offset(-1), rg.End(xlToRight).End(xlDown)).AutoFilter
    On Error GoTo ERR_MISSING_FILTER
    ws.AutoFilter.sort.SortFields.clear
    On Error GoTo 0
    ws.AutoFilter.sort.SortFields.add Key:= _
        Range( _
            rg.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED), _
            rg.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).End(xlDown) _
            ), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal

    With ws.AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
ERR_MISSING_FILTER:
    ws.Range(rg.Offset(-1), rg.End(xlToRight).End(xlDown)).AutoFilter
End Function


Public Function sort_descending(wb As Workbook)
    Dim ws As Worksheet
    Dim rg As Range
    'Dim rg_sort As Range
    
    Set ws = wb.Worksheets(STR_WS_NAME)
    Set rg = ws.Range(STR_DATA_FIRST_CELL)
    
    'ws_local.Activate
    'ws.Range(rg.Offset(-1), rg.End(xlToRight).End(xlDown)).AutoFilter
    On Error GoTo ERR_MISSING_FILTER
    ws.AutoFilter.sort.SortFields.clear
    On Error GoTo 0
    ws.AutoFilter.sort.SortFields.add Key:= _
        Range( _
            rg.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED), _
            rg.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).End(xlDown) _
            ), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal

    With ws.AutoFilter.sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
ERR_MISSING_FILTER:
    ws.Range(rg.Offset(-1), rg.End(xlToRight).End(xlDown)).AutoFilter
End Function

