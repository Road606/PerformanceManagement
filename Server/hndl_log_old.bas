Attribute VB_Name = "hndl_log_old"
Option Explicit

' file handling
Public STR_PATH_INBOUND As String
Public str_file_name As String

Public STR_WS_NAME As String
Public STR_DATA_FIRST_ROW_RG As String
Public STR_DATA_START_RG As String

Public obj_log_record As DBLogRecord
Public wb As Workbook
Public ws As Worksheet

Public Function init()
    STR_WS_NAME = "db.log"
    STR_DATA_FIRST_ROW_RG = "A2:E2"
    STR_DATA_START_RG = "A2"
    
    db_log.init
        
End Function

Public Function open_data()
    Set wb = util_file.open_wb(STR_PATH_INBOUND & str_file_name, False)
    Set ws = wb.Worksheets(STR_WS_NAME)
End Function

Public Function close_data()
    wb.Close SaveChanges:=True
End Function

Public Function get_data() As Range
    Set get_data = ws.Range(STR_DATA_FIRST_ROW_RG)
    
    Set get_data = ws.Range( _
        get_data, get_data.End(xlDown))
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
