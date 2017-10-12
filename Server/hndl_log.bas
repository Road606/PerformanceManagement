Attribute VB_Name = "hndl_log"
Option Explicit

' file handling
Public str_path As String
Public str_file_name As String

Public STR_WS_NAME As String
Public STR_DATA_FIRST_ROW_RG As String
Public STR_DATA_START_RG As String

Public BOOL_EXTERNAL_DATA_FILE_VISIBILITY As Boolean

Public obj_log_record As DBLogRecord
Public wb As Workbook
Public ws As Worksheet

Public lng_logs As Long
Public lng_err_logs As Long
Public lng_warn_logs As Long
Public lng_info_logs As Long

Public Property Get file_path() As String
    file_path = str_path & str_file_name
End Property

Public Function init()
    STR_WS_NAME = "db.log"
    STR_DATA_FIRST_ROW_RG = "A2:E2"
    STR_DATA_START_RG = "A2"
    
    BOOL_EXTERNAL_DATA_FILE_VISIBILITY = False
    
    lng_logs = 0
    lng_err_logs = 0
    lng_warn_logs = 0
    lng_info_logs = 0
    
    db_log.init
End Function

Public Function open_data()
    'Set wb = util_file.open_wb(STR_PATH & STR_FILE_NAME, False)
    Set wb = util_file.open_wb(file_path, False, BOOL_EXTERNAL_DATA_FILE_VISIBILITY)
    Set ws = wb.Worksheets(STR_WS_NAME)
End Function

Public Function close_data()
    Windows(wb.Name).Visible = True
    wb.Close SaveChanges:=True
End Function

Public Function get_data() As Range
    Set get_data = ws.Range(STR_DATA_FIRST_ROW_RG)
    
    Set get_data = ws.Range( _
        get_data, get_data.End(xlDown))
End Function

Public Function log(str_type As String, str_module As String, str_function As String, str_message As String)
    Set obj_log_record = New DBLogRecord
    
    obj_log_record.str_datetime = Now
    obj_log_record.str_type = str_type
    obj_log_record.str_module = str_module
    obj_log_record.str_function = str_function
    obj_log_record.str_message = str_message
    
    save_record
    
    lng_logs = lng_logs + 1
End Function

Public Function log_err(str_module As String, str_function As String, str_message As String)
    log db_log.TYPE_ERR, str_module, str_function, str_message
End Function

Public Function log_info(str_module As String, str_function As String, str_message As String)
    log db_log.TYPE_INFO, str_module, str_function, str_message
End Function

Public Function log_warn(str_module As String, str_function As String, str_message As String)
    log db_log.TYPE_WARN, str_module, str_function, str_message
End Function

Public Function save_record()
    Dim rg As Range
    
    On Error GoTo ERR_FULL_SHEET
    Set rg = next_row()
    rg.Offset(0, db_log.INT_DATA_COL_OFFSET_DATETIME).Value = obj_log_record.str_datetime
    rg.Offset(0, db_log.INT_DATA_COL_OFFSET_TYPE).Value = obj_log_record.str_type
    rg.Offset(0, db_log.INT_DATA_COL_OFFSET_MODULE).Value = obj_log_record.str_module
    rg.Offset(0, db_log.INT_DATA_COL_OFFSET_FUNCTION).Value = obj_log_record.str_function
    rg.Offset(0, db_log.INT_DATA_COL_OFFSET_MESSAGE).Value = obj_log_record.str_message
    On Error GoTo 0
    Exit Function
ERR_FULL_SHEET:
    MsgBox "Log file is full.", vbCritical, "Log error"
End Function

Public Function next_row() As Range
    Set next_row = ws.Cells(ws.Range("A:A").CountLarge, 1).End(xlUp).Offset(1)
End Function

Public Function is_err_level() As Boolean
    is_err_level = lng_err_logs > 0
End Function

Public Function is_warn_level() As Boolean
    is_warn_level = lng_warn_logs > 0
End Function

Public Function is_info_level() As Boolean
    is_info_level = lng_info_logs > 0
End Function

Public Function reset_counters()
    lng_logs = 0
    lng_err_logs = 0
    lng_warn_logs = 0
    lng_info_logs = 0
End Function
