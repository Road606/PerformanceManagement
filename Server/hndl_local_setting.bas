Attribute VB_Name = "hndl_local_setting"
Option Explicit

' local data
Public Const STR_LOCAL_WS_NAME = "local.setting"
Public Const STR_LOCAL_DATA_START_RG = "A2"

' raw data
Public Const STR_FIRST_ROW_RG As String = "A2:B2"

  ' columns
Public Const STR_LOCAL_DATA_COL_OFFSET_NAME As String = "0"
Public Const STR_LOCAL_DATA_COL_OFFSET_VALUE As String = "1"

Public Function init()
End Function

Public Function get_value(str_name As String) As String
    Dim rg_data As Range
    
    Set rg_data = get_data
    'Debug.Print rg_data.Address
    get_value = WorksheetFunction.VLookup( _
            str_name, _
            rg_data, _
            CInt(STR_LOCAL_DATA_COL_OFFSET_VALUE) - CInt(STR_LOCAL_DATA_COL_OFFSET_NAME) + 1, _
            False _
        )
End Function

Public Function get_data() As Range
    Set get_data = ThisWorkbook.Worksheets(STR_LOCAL_WS_NAME).Range(STR_FIRST_ROW_RG)
    Set get_data = Range(get_data, get_data.End(xlDown))
End Function

