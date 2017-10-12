VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiFileExcelDataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MultiFileExcelDataProvider"

' providers management
  ' general
Public str_source_type As String
  ' file
Public str_path As String
Public str_file_appendix As String
Public str_tmpl_file_path As String
Public bool_visibility_mode As Boolean
Public str_passwd As String
  ' sheet
Public STR_WS_NAME As String
Public STR_DATA_FIRST_CELL As String
  ' providers
Public col_providers As Collection
    ' collaboration
      ' listeners must implement:
        ' loading_data_is_about_to_start obj_provider_info:FileExcelDataProviderInfo
Public col_listeners As Collection
    ' action
      ' loading
Public BOOL_IS_DYNAMIC_LOADING_PROVIDERS_FROM_FILES_ON As Boolean

' tracking management
Public STR_PROCESSED_FILE_PATH As String
Public STR_PROCESSED_FILE_NAME As String

' status bar
'Public STR_STATUS_BAR_PREFIX As String
'Public STR_STATUS_BAR_PREFIX_LOADING As String
'Public STR_STATUS_BAR_PREFIX_LOADING_FINISHED As String
'Public LNG_STATUS_BAR_REFRESH As Long = 100

Public obj_current_data_provider_info As FileExcelDataProviderInfo

Public obj_mdl_processed_file As MDLProcessedFile
Public obj_data_provider_util As FileExcelDataProviderUtil

Private Sub Class_Initialize()
    ' providers management
      ' file
    STR_DATA_FIRST_CELL = "A2"
    bool_visibility_mode = new_const_excel_data_provider.BOOL_VISIBILITY_MODE_HIDDEN
      ' providers
    Set col_providers = New Collection
        ' collaboration
    Set col_listeners = New Collection
    
    ' action
      ' loading
    BOOL_IS_DYNAMIC_LOADING_PROVIDERS_FROM_FILES_ON = False
End Sub

' action
Public Function load_data_static()
    Dim obj_data_provider As FileExcelDataProvider

    For Each obj_data_provider In col_providers
        On Error GoTo ERR_PROCESS_FILE
        before_load_data_static obj_data_provider
        obj_data_provider.load_data
        after_load_data_static
        On Error GoTo 0
    Next

    Exit Function
ERR_PROCESS_FILE:
    hndl_log.log db_log.TYPE_WARN, str_module, "process", _
        "An error occured during processing file " & obj_data_provider.str_id & ". Original error message>" & Err.description
    Resume Next
End Function


Public Function load_data()
    Dim obj_data_provider As FileExcelDataProvider
    
    On Error GoTo ERR_PROCESSED_FILE_NOT_FOUND
    Set obj_mdl_processed_file = New MDLProcessedFile
    obj_mdl_processed_file.str_path = STR_PROCESSED_FILE_PATH
    obj_mdl_processed_file.str_file_name = STR_PROCESSED_FILE_NAME
    obj_mdl_processed_file.open_data
    On Error GoTo 0
    
    On Error GoTo ERR_CREATE_PROVIDERS
    If BOOL_IS_DYNAMIC_LOADING_PROVIDERS_FROM_FILES_ON Then
        create_data_providers_read_mode
    End If
    On Error GoTo 0
    
    For Each obj_data_provider In col_providers
        On Error GoTo ERR_PROCESS_FILE
        before_load_data obj_data_provider
        obj_data_provider.load_data 'obj_current_data_provider_info
        after_load_data
        remove_provider obj_data_provider
        On Error GoTo 0
    Next
            
    ' clear
    obj_mdl_processed_file.close_data
            
    Exit Function
ERR_PROCESSED_FILE_NOT_FOUND:
    hndl_log.log db_log.TYPE_WARN, str_module, "process", Err.Number & "->" & Err.description
    Exit Function
ERR_CREATE_PROVIDERS:
    hndl_log.log db_log.TYPE_WARN, str_module, "process", Err.Number & "->" & Err.description
    Exit Function
ERR_PROCESS_FILE:
    hndl_log.log db_log.TYPE_WARN, str_module, "process", _
        "An error occured during processing file " & obj_data_provider.str_id & ". Original error message>" & Err.description
    Resume Next
End Function

Public Function before_load_data(obj_provider As FileExcelDataProvider)
    Dim obj_listener As Object

    ' register processed file
      ' # implement - processed_file will be registered as usual listener
    Set obj_mdl_processed_file.obj_file_record = New DBProcessedFileRecord
    obj_mdl_processed_file.obj_file_record.str_name = obj_provider.str_file_name
    obj_mdl_processed_file.obj_file_record.str_date_started = Now
    ' send basic data about current processed provider
    Set obj_current_data_provider_info = New FileExcelDataProviderInfo
        obj_current_data_provider_info.str_type = new_const_excel_data_provider.STR_DATA_PROVIDER_INFO_DYNAMIC
        obj_current_data_provider_info.str_provider_id = obj_provider.str_id
        Set obj_current_data_provider_info.obj_period = _
            obj_data_provider_util.retrieve_period( _
                obj_data_provider_util.retrieve_provider_id_reverse(obj_provider.str_file_name))
                
    For Each obj_listener In col_listeners
        obj_listener.loading_data_is_about_to_start obj_current_data_provider_info
    Next
    
End Function

Public Function after_load_data()
    Dim obj_listener As Object
    Dim obj_provider_info As FileExcelDataProviderInfo

    ' # implement - processed_file will be registered as usual listener
    obj_mdl_processed_file.obj_file_record.str_date_finished = Now
    obj_mdl_processed_file.save_record
    
    For Each obj_listener In col_listeners
        obj_listener.loading_data_has_finished obj_current_data_provider_info
    Next
    
    Set obj_current_data_provider_info = Nothing
End Function

Public Function before_load_data_static(obj_provider As FileExcelDataProvider)
    Dim obj_listener As Object
    
    Set obj_current_data_provider_info = New FileExcelDataProviderInfo
    obj_current_data_provider_info.str_type = new_const_excel_data_provider.STR_DATA_PROVIDER_INFO_STATIC
    obj_current_data_provider_info.str_provider_id = obj_provider.str_id
    
    For Each obj_listener In col_listeners
        obj_listener.loading_data_is_about_to_start obj_current_data_provider_info
    Next
End Function

Public Function after_load_data_static()
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
        obj_listener.loading_data_has_finished obj_current_data_provider_info
    Next
    
    Set obj_current_data_provider_info = Nothing
End Function

' providers management
Public Function create_data_providers_read_mode()
    Dim obj_list_files As Collection
    Dim var_file As Variant
    Dim obj_data_provider As FileExcelDataProvider
    
    On Error GoTo ERR_RETRIEVE_FILES
    Set obj_list_files = retrieve_files
    On Error GoTo 0
    
    For Each var_file In obj_list_files
        On Error GoTo ERR_PROCESS_FILE
        Set obj_data_provider = create_data_provider_read_mode( _
            obj_data_provider_util.retrieve_provider_id_reverse(CStr(var_file)), CStr(var_file), str_source_type)
        On Error GoTo 0
    Next
    Exit Function
ERR_RETRIEVE_FILES:
    hndl_log.log db_log.TYPE_WARN, str_module, "process", Err.Number & "->" & Err.description
    Exit Function
ERR_PROCESS_FILE:
    hndl_log.log db_log.TYPE_WARN, str_module, "process", _
        "An error occured during processing file " & var_file & ". Original error message>" & Err.description
    Exit Function
End Function

Private Function retrieve_files() As Collection
    Dim rg_processed_file As Range
    
    Set retrieve_files = util_file.retrieve_files(str_path, "*" & str_file_appendix)
    Set rg_processed_file = obj_mdl_processed_file.ws.Range(obj_mdl_processed_file.STR_DATA_START_RG)
    
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

Public Function create_data_provider_save_mode(str_id As String, Optional str_file_name As String, Optional str_tmpl_file_path_unique As String, Optional str_source_type_unique As String = "") As FileExcelDataProvider
    Dim str_tmpl_file_path_final As String
    
    If str_tmpl_file_path_unique <> "" Then
        str_tmpl_file_path_final = str_tmpl_file_path_unique
    Else
        str_tmpl_file_path_final = str_tmpl_file_path
    End If
    
    Set create_data_provider_save_mode = create_data_provider(str_id, str_file_name, str_source_type_unique)
    create_data_provider_save_mode.str_tmpl_file_path = str_tmpl_file_path_final
    create_data_provider_save_mode.bool_save_mode_on = True
End Function

Public Function create_data_provider_read_mode(str_id As String, Optional str_file_name As String, Optional str_source_type_unique As String = "") As FileExcelDataProvider
    Set create_data_provider_read_mode = create_data_provider(str_id, str_file_name, str_source_type_unique)
    create_data_provider_read_mode.bool_save_mode_on = False
End Function

Private Function create_data_provider(str_id As String, Optional str_file_name As String, Optional str_source_type_unique As String = "") As FileExcelDataProvider
    Dim obj_listener As Object
    Dim str_source_type_final
    
    Set create_data_provider = New FileExcelDataProvider
    create_data_provider.str_id = str_id
    
    If str_source_type_unique <> "" Then
        str_source_type_final = str_source_type_unique
    Else
        str_source_type_final = str_source_type
    End If
    
    ' general
    create_data_provider.str_source_type = str_source_type_final
    
    If str_source_type_final = new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL Then
        ' file
        create_data_provider.str_path = str_path
        create_data_provider.str_file_name = str_file_name
        create_data_provider.bool_visibility_mode = bool_visibility_mode
    End If
    
    ' sheet
    create_data_provider.STR_WS_NAME = STR_WS_NAME
    ' collaboration
    For Each obj_listener In col_listeners
        create_data_provider.add_listener obj_listener
    Next
    
    add_provider create_data_provider
End Function

Public Function add_provider(obj_data_provider As FileExcelDataProvider)
    col_providers.add obj_data_provider, obj_data_provider.str_id
End Function

Public Function retrieve_provider(str_provider_id As String) As FileExcelDataProvider
    Set retrieve_provider = col_providers(str_provider_id)
End Function

Public Function remove_provider(obj_data_provider As FileExcelDataProvider)
    col_providers.Remove obj_data_provider.str_id
End Function

Public Function close_providers()
    Dim obj_provider As FileExcelDataProvider
    
    On Error GoTo WARN_CLOSING_PROVIDER_FAILED
    For Each obj_provider In col_providers
        obj_provider.close_data
        remove_provider obj_provider
    Next
    On Error GoTo 0
    Exit Function
WARN_CLOSING_PROVIDER_FAILED:
    hndl_log.log db_log.TYPE_WARN, str_module, "close_providers", _
        "Closing provider: " & obj_provider.str_id & " has failed."
End Function

' automation
Public Function set_clear_data_after_open_status(bool_yes_no As Boolean)
    Dim obj_provider As FileExcelDataProvider
    
    For Each obj_provider In col_providers
        obj_provider.set_clear_data_after_open_status bool_yes_no
    Next
End Function

Public Function set_clear_data_before_close_status(bool_yes_no As Boolean)
    Dim obj_provider As FileExcelDataProvider
    
    For Each obj_provider In col_providers
        obj_provider.set_clear_data_before_close_status bool_yes_no
    Next
End Function

Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function
