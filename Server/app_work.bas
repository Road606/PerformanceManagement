Attribute VB_Name = "app_work"
Option Explicit

' local data
Public Const STR_DATA_FIRST_ROW_RG As String = "A2:CR2"
Public Const STR_WS_NAME = "app.work"
'Public Const STR_LOCAL_WS_NAME = "Sheet2"
Public Const STR_DATA_START_RG = "A2"

Public Function clear()
    ThisWorkbook.Worksheets(STR_WS_NAME).UsedRange.ClearContents
End Function
