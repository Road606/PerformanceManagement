Attribute VB_Name = "app_status_bar"
Option Explicit

Public lng_refresh_num As Long
Public lng_total_records As Long

Public str_module As String
Public str_method As String
Public STR_FILE As String

Public Function update_records(lng_record_num As Long)
    If lng_record_num Mod lng_refresh_num = 0 Then
            Application.StatusBar = str_module & str_method & "from file: " & STR_FILE & ". " & (WorksheetFunction.RoundUp(lng_record_num / lng_total_records, 2) * 100) & "% of records has been processed."
        End If
End Function
