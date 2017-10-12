Attribute VB_Name = "new_file_processed_level1"
Option Explicit

' file handling
Public STR_PATH_INBOUND As String
Public str_file_name As String

Public STR_WS_NAME As String
Public STR_DATA_FIRST_ROW_RG As String
Public STR_DATA_START_RG As String

Public obj_file_record As DBProcessedFileRecord
Public wb As Workbook
Public ws As Worksheet

Public Function init()
    STR_WS_NAME = "db.file.processed"
    STR_DATA_FIRST_ROW_RG = "A2:CR2"
    STR_DATA_START_RG = "A2"
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
    rg.Offset(0, new_db_file_processed.INT_DATA_COL_OFFSET_NAME).Value = obj_file_record.str_name
    rg.Offset(0, new_db_file_processed.INT_DATA_COL_OFFSET_TRANSACTION_STARTED).Value = obj_file_record.str_date_started
    rg.Offset(0, new_db_file_processed.INT_DATA_COL_OFFSET_TRANSACTION_FINISHED).Value = obj_file_record.str_date_finished
    On Error GoTo 0
    Exit Function
ERR_FULL_SHEET:
    MsgBox "Log of processed files is full.", vbCritical, "Log error"
End Function

Public Function next_row() As Range
    Set next_row = ws.Cells(ws.Range("A:A").CountLarge, 1).End(xlUp).Offset(1)
End Function


