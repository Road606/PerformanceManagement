VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileExcelDataProviderStatusUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_module As String

Public str_provider_id As String
Public str_provider_status As String

' date time
Private private_str_datetime As String

' save mode
Public str_save_mode As String

' file name
  ' tmpl
Public str_tmpl_path As String
Public str_tmpl_file_name_prefix As String
Public str_tmpl_file_name_appendix As String
  ' new
Public str_file_prefix As String
Public str_file_name_separator As String
Public str_file_appendix As String

' regex
Public str_regex_pattern As String
Public int_regex_id As Integer
Public int_regex_status As Integer

Public Property Get str_datetime() As String
    str_datetime = private_str_datetime
End Property

Public Property Let str_datetime(str_datetime As String)
    private_str_datetime = str_datetime
    update_provider_id
End Property

Public Function create_provider_info(str_file_name As String, str_type As String, str_loading_type As String) As FileExcelDataProviderInfo
    Set create_provider_info = New FileExcelDataProviderInfo
    create_provider_info.str_provider_id = retrieve_provider_id_reverse(str_file_name)
    Set create_provider_info.obj_period = retrieve_period(create_provider_info.str_provider_id)
    create_provider_info.str_type = str_type
    create_provider_info.str_loading_type = str_loading_type
    create_provider_info.str_completeness_status = retrieve_provider_status_reverse(str_file_name)
End Function

Public Function retrieve_creation_tmpl_file_name()
    retrieve_creation_tmpl_file_name = _
        str_tmpl_path & _
        str_tmpl_file_name_prefix & _
        str_file_name_separator & _
        str_tmpl_file_name_appendix & _
        str_file_appendix
End Function

'Public Function retrieve_creation_file_name()
'    retrieve_creation_file_name = _
'        str_file_prefix & _
'        str_file_name_separator & _
'        str_provider_id & _
'        str_file_appendix
'End Function

Public Function retrieve_creation_file_name()
    retrieve_creation_file_name = _
        str_file_prefix & _
        str_file_name_separator & _
        str_provider_id & _
        str_file_name_separator & _
        str_provider_status & _
        str_file_appendix
End Function

Public Function retrieve_period(str_provider_id As String) As Period
    Dim str_year As String
    Dim str_month As String
    Dim str_day As String
    Dim str_hour As String
    Dim str_minute As String

    Set retrieve_period = New Period

    Select Case str_save_mode
        Case new_const_excel_data_provider.STR_SAVE_MODE_HOURLY
            str_year = Left(str_provider_id, 2)
            str_month = Mid(str_provider_id, 3, 2)
            str_day = Mid(str_provider_id, 5, 2)
            str_hour = Mid(str_provider_id, 7, 2)
            
            retrieve_period.str_start = DateValue(str_day & "." & str_month & "." & str_year) + TimeValue(str_hour & ":00:00")
            retrieve_period.str_end = DateAdd("h", 1, retrieve_period.str_start)
        Case new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_MINUTE
            str_year = Left(str_provider_id, 2)
            str_month = Mid(str_provider_id, 3, 2)
            str_day = Mid(str_provider_id, 5, 2)
            str_hour = Mid(str_provider_id, 7, 2)
            str_minute = Mid(str_provider_id, 9, 2)
            
            retrieve_period.str_start = DateValue(str_day & "." & str_month & "." & str_year) + TimeValue(str_hour & ":" & str_minute & ":00")
            retrieve_period.str_end = DateAdd("h", 1, retrieve_period.str_start)
    End Select
End Function

'Public Function retrieve_provider_id_reverse(str_file_name) As String
'    Dim str_format As String
'
'    Select Case str_save_mode
'        Case new_const_excel_data_provider.STR_SAVE_MODE_HOURLY
'            str_format = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_FORMAT
'        Case new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_MINUTE
'            str_format = new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_MINUTE_FORMAT
'    End Select
'
'    retrieve_provider_id_reverse = _
'        Mid(str_file_name, _
'            Len(str_file_prefix & str_file_name_separator) + 1, _
'            Len(str_format))
'End Function

Public Function retrieve_provider_id_reverse(str_file_name) As String
    Dim regex As Object
    Dim match As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = str_regex_pattern
    
    On Error GoTo ERR_RETRIEVAL_FAILED
    Set match = regex.Execute(str_file_name)
    retrieve_provider_id_reverse = CStr(match(0).SubMatches(int_regex_id))
    On Error GoTo 0
    Exit Function
ERR_RETRIEVAL_FAILED:
    hndl_log.log db_log.TYPE_WARN, str_module, "retrieve_provider_id", "Retrieval of provider has failed. Pattern: " & str_regex_pattern & ", Filename: " & str_file_name & "."
End Function

Public Function retrieve_provider_status_reverse(str_file_name) As String
    Dim regex As Object
    Dim match As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = str_regex_pattern
    
    Set match = regex.Execute(str_file_name)
    retrieve_provider_status_reverse = CStr(match(0).SubMatches(int_regex_status))
End Function

Private Function update_provider_id()
    Select Case str_save_mode
        Case new_const_excel_data_provider.STR_SAVE_MODE_HOURLY
            str_provider_id = Format(private_str_datetime, new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_FORMAT)
        Case new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_MINUTE
            str_provider_id = Format(private_str_datetime, new_const_excel_data_provider.STR_SAVE_MODE_HOURLY_MINUTE_FORMAT)
    End Select
End Function


Private Sub Class_Initialize()
    str_module = "FileExcelDataProviderStatusUtil"
End Sub
