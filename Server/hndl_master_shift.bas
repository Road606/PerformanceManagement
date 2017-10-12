Attribute VB_Name = "hndl_master_shift"
Option Explicit

' local data
Public Const STR_LOCAL_WS_NAME = "data.master.shift"
Public Const STR_LOCAL_DATA_START_RG = "A2"

' raw data
Public Const STR_FIRST_ROW_RG As String = "A2:D2"

  ' columns
Public Const STR_LOCAL_DATA_COL_OFFSET_ID As String = "0"
Public Const STR_LOCAL_DATA_COL_OFFSET_SHIFT As String = "3"

Public obj_missing_data As Collection

Public Function init()
    Set obj_missing_data = New Collection
End Function

Public Function find_shift(str_id As String) As String
    Dim rg_data As Range
    Dim lng_qty_unit As Long
    Dim lng_qty_stock As Long
    
    Set rg_data = get_data
    'Debug.Print rg_data.Address
    find_shift = WorksheetFunction.VLookup( _
            str_id, _
            rg_data, _
            CInt(STR_LOCAL_DATA_COL_OFFSET_SHIFT) - CInt(STR_LOCAL_DATA_COL_OFFSET_ID) + 1, _
            False _
        )
End Function

Public Function find_type(int_hour As Integer) As String
    If int_hour > 6 And int_hour < 18 Then
        find_type = "D"
    Else
        find_type = "N"
    End If
End Function

Public Function get_data() As Range
    Set get_data = ThisWorkbook.Worksheets(STR_LOCAL_WS_NAME).Range(STR_FIRST_ROW_RG)
    Set get_data = Range(get_data, get_data.End(xlDown))
End Function


