VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLBINProdLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLBINProdLine"

Public STR_DB_TYPE_FILE_EXCEL As String

Public INT_OFFSET_PROD_LINE_ID As Integer
Public INT_OFFSET_WC As Integer
Public INT_OFFSET_BUILDING As Integer

'Public obj_multi_data_provider As MultiFileExcelDataProvider
Public obj_data_provider As FileExcelDataProvider
Public bool_load_into_local_collection As Boolean

'Public obj_current_data_provider_info As FileExcelDataProviderInfo
'Public str_static_data_provider_id As String
'Public str_provider_id_default As String

Public col_bin_prod_lines As Collection
    ' listeners must implement:
      ' process_record obj_db_data_process: DBDataProcess
Public col_listeners As Collection
    ' listeners must implement:
      ' loading_data_is_about_to_start obj_provider_info: FileExcelDataProviderInfo
'Public col_provider_info_listeners As Collection

Private Sub Class_Initialize()
    STR_DB_TYPE_FILE_EXCEL = "file:excel"

    INT_OFFSET_PROD_LINE_ID = 0
    INT_OFFSET_WC = 1
    INT_OFFSET_BUILDING = 2
    
    'Set obj_multi_data_provider = New MultiFileExcelDataProvider
    'str_static_data_provider_id = "static"
    Set col_listeners = New Collection
    Set col_bin_prod_lines = New Collection
End Sub

Public Function load()
    obj_data_provider.add_listener Me
    obj_data_provider.load_data
End Function

Public Function load_local()
    obj_data_provider.add_listener Me
    obj_data_provider.load_data
End Function

' # interface method for listening FileExcelDataProvider
Public Function load_record(rg_record As Range)
    Dim obj_bin_prod_line As DBBINProdLine
    Dim obj_listener As Object
    Dim errorMsg As MSG

    ' fill data in a row
    Set obj_bin_prod_line = New DBBINProdLine
      ' pallet level
    obj_bin_prod_line.str_prod_line_id = rg_record.Offset(0, INT_OFFSET_PROD_LINE_ID).Value
    obj_bin_prod_line.str_wc = rg_record.Offset(0, INT_OFFSET_WC).Value
    obj_bin_prod_line.str_building = rg_record.Offset(0, INT_OFFSET_BUILDING).Value
    
    If bool_load_into_local_collection Then
        col_bin_prod_lines.add obj_bin_prod_line, obj_bin_prod_line.str_prod_line_id
    End If
    
    For Each obj_listener In col_listeners
        obj_listener.process_record obj_bin_prod_line
    Next
    Exit Function
WARN_STOCK_UNIT:
    Set errorMsg = New MSG
    log4VBA.warn log4VBA.DEFAULT_DESTINATION, errorMsg.source(STR_APP_MODULE_NAME, "load_record").text("Invalid value for Stock unit on row: " & rg_record.Row)
    hndl_log.log db_log.TYPE_WARN, STR_APP_MODULE_NAME, "load_record", "Invalid value for Stock unit on row: " & rg_record.Row
    Resume Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object
    Dim errorMsg As MSG

    'Set obj_current_data_provider_info = obj_provider_info
    
    For Each obj_listener In col_listeners
        On Error GoTo WARN_LISTENER_ERR
        obj_listener.loading_data_is_about_to_start obj_provider_info
        On Error GoTo 0
    Next
    
    Exit Function
WARN_LISTENER_ERR:
    Set errorMsg = New MSG
    log4VBA.warn log4VBA.DEFAULT_DESTINATION, errorMsg.source(STR_APP_MODULE_NAME, "loading_data_is_about_to_start").text("Something went wrong for listener: " & obj_listener.STR_CLASS_NAME)
    'hndl_log.log db_log.TYPE_WARN, STR_APP_MODULE_NAME, "loading_data_has_finished", "Something went wrong for listener: " & obj_listener.STR_CLASS_NAME
    Resume Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object
    Dim errorMsg As MSG
    
    For Each obj_listener In col_listeners
        On Error GoTo WARN_LISTENER_ERR
        obj_listener.loading_data_has_finished obj_provider_info
        On Error GoTo 0
    Next
    
    'Set obj_current_data_provider_info = Nothing
    
    Exit Function
WARN_LISTENER_ERR:
    Set errorMsg = New MSG
    log4VBA.warn log4VBA.DEFAULT_DESTINATION, errorMsg.source(STR_APP_MODULE_NAME, "loading_data_has_finished").text("Something went wrong for listener: " & obj_listener.STR_CLASS_NAME)
    'hndl_log.log db_log.TYPE_WARN, STR_APP_MODULE_NAME, "loading_data_has_finished", "Something went wrong for listener: " & obj_listener.STR_CLASS_NAME
    Resume Next
End Function

'Public Function save_record_dynamic(obj_data_process_record As DBDataProcess)
'    Dim obj_data_provider As FileExcelDataProvider
'    Dim obj_data_provider_util As FileExcelDataProviderUtil
'    Dim str_provider_id As String
'
'    Set obj_data_provider_util = obj_multi_data_provider.obj_data_provider_util
'    If obj_data_process_record.str_date_end = "" Then
'        obj_data_provider_util.str_datetime = obj_data_process_record.str_date_start
'    Else
'        obj_data_provider_util.str_datetime = obj_data_process_record.str_date_end 'obj_data_process_record.str_date_start
'    End If
'
'    str_provider_id = obj_data_provider_util.str_provider_id 'retrieve_provider_id(obj_data_process_record)
'
'    On Error GoTo INFO_PROVIDER_NOT_FOUND
'    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(str_provider_id)
'    On Error GoTo 0
'
'    save_record obj_data_process_record, obj_data_provider.next_data
'    Exit Function
'INFO_PROVIDER_NOT_FOUND:
'    Set obj_data_provider = _
'        obj_multi_data_provider.create_data_provider_save_mode( _
'            str_provider_id, _
'            obj_data_provider_util.retrieve_creation_file_name, _
'            obj_data_provider_util.retrieve_creation_tmpl_file_name)
'    obj_data_provider.open_data
'    Resume Next
'End Function
'
'Public Function save_record_static(obj_data_process_record As DBDataProcess)
'    Dim rg_next_data As Range
'
'    On Error GoTo INFO_DATA_PROVIDER_NOT_OPEN_YET
'    Set rg_next_data = obj_multi_data_provider.retrieve_provider(str_static_data_provider_id).next_data
'    On Error GoTo 0
'
'    save_record obj_data_process_record, rg_next_data
'    Exit Function
'INFO_DATA_PROVIDER_NOT_OPEN_YET:
'    obj_multi_data_provider.retrieve_provider(str_static_data_provider_id).open_data
'    Set rg_next_data = obj_multi_data_provider.retrieve_provider(str_static_data_provider_id).next_data
'    Resume Next
'End Function
'
'Public Function save_record(obj_data_process_record As DBDataProcess, rg_record As Range)
'
'    rg_record.Offset(0, INT_OFFSET_PALLET).NumberFormat = "@"
'    rg_record.Offset(0, INT_OFFSET_PALLET).Value = obj_data_process_record.str_pallet
'    rg_record.Offset(0, INT_OFFSET_MATERIAL).Value = obj_data_process_record.str_material
'    rg_record.Offset(0, INT_OFFSET_MATERIAL_VENDOR).Value = obj_data_process_record.str_material_vendor
'    rg_record.Offset(0, INT_OFFSET_MATERIAL_BIN_STORAGE_GROUP).Value = obj_data_process_record.str_material_bin_storage_group
'    rg_record.Offset(0, INT_OFFSET_SOURCE).Value = obj_data_process_record.str_source
'    rg_record.Offset(0, INT_OFFSET_SOURCE_TYPE).Value = obj_data_process_record.str_source_type
'    rg_record.Offset(0, INT_OFFSET_DESTINATION).Value = obj_data_process_record.str_destination
'    rg_record.Offset(0, INT_OFFSET_DESTINATION_TYPE).Value = obj_data_process_record.str_destination_type
'    rg_record.Offset(0, INT_OFFSET_STOCK_UNIT).Value = obj_data_process_record.lng_stock_unit
'      ' process level
'    rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value = obj_data_process_record.str_creation_id
'    rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION).Value = obj_data_process_record.str_version
'      ' step level
'    rg_record.Offset(0, INT_OFFSET_USER).Value = obj_data_process_record.str_user
'    rg_record.Offset(0, INT_OFFSET_DATE_TIME_START).Value = obj_data_process_record.str_date_start
'    rg_record.Offset(0, INT_OFFSET_DATE_TIME_END).Value = obj_data_process_record.str_date_end
'    rg_record.Offset(0, INT_OFFSET_BIN_START).Value = obj_data_process_record.str_bin_from
'    rg_record.Offset(0, INT_OFFSET_BIN_END).Value = obj_data_process_record.str_bin_to
'    rg_record.Offset(0, INT_OFFSET_TRANSACTION_TYPE_START).Value = obj_data_process_record.str_transaction_type_start
'    rg_record.Offset(0, INT_OFFSET_TRANSACTION_TYPE_END).Value = obj_data_process_record.str_transaction_type_end
'    'rg_record.Offset(0, INT_OFFSET_PROCESS_STEP).Value = obj_process_step.obj_master_step.
'    rg_record.Offset(0, INT_OFFSET_PROCESS_STATUS).Value = obj_data_process_record.byte_process_status
'End Function

Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function

'Public Function add_provider_info_listener(obj_listener As Object)
'    col_provider_info_listeners.add obj_listener
'End Function

'Public Function retrieve_provider_id(obj_data_process_record As DBDataProcess) As String
'    Dim str_date_identifier As String
'
'    Select Case STR_SAVE_MODE
'        Case STR_SAVE_MODE_HOURLY
'            retrieve_provider_id = Format(obj_data_process_record.str_date_start, "YYMMDDHH")
'        Case STR_SAVE_MODE_DAILY
'            retrieve_provider_id = Format(obj_data_process_record.str_date_start, "YYMMDD")
'    End Select
'End Function
'
'Public Function retrieve_creation_tmpl_file_name()
'    retrieve_creation_tmpl_file_name = _
'        STR_CREATION_TMPL_PATH & _
'        STR_CREATION_FILE_PREFIX & _
'        STR_CREATION_FILE_NAME_SEPARATOR & _
'        STR_CREATION_TMPL_FILE_APPENDIX & _
'        STR_CREATION_FILE_APPENDIX
'End Function
'
'Public Function retrieve_creation_file_name(str_date_identifier As String)
'    retrieve_creation_file_name = _
'        STR_CREATION_FILE_PREFIX & _
'        STR_CREATION_FILE_NAME_SEPARATOR & _
'        str_date_identifier & _
'        STR_CREATION_FILE_APPENDIX
'End Function

'Public Function retrieve_creation_file_name(str_date_identifier As String)
'    retrieve_creation_file_name = _
'        STR_CREATION_PATH & _
'        STR_CREATION_FILE_PREFIX & _
'        STR_CREATION_FILE_NAME_SEPARATOR & _
'        str_date_identifier & _
'        STR_CREATION_FILE_APPENDIX
'End Function
'Public Function retrieve_dates() As Collection
'    Dim rg_app_work As Range
'    Dim rg_transaction_date As Range
'
'    app_work.clear
'
'    Set rg_transaction_date = _
'        ThisWorkbook.Worksheets(hndl_performance.str_ws_name). _
'        Range(hndl_performance.STR_DATA_FIRST_CELL). _
'        Offset(0, db_performance.INT_OFFSET_TRANSACTION_DATE)
'
'    Set rg_app_work = _
'            ThisWorkbook.Worksheets(app_work.str_ws_name). _
'            Range(app_work.STR_DATA_START_RG)
'
'    If rg_transaction_date.Value <> "" Then
'        Set rg_transaction_date = Range(rg_transaction_date, rg_transaction_date.End(xlDown))
'        rg_transaction_date.Copy
'
'        rg_app_work.PasteSpecial xlPasteValues
'        rg_app_work.RemoveDuplicates Columns:=Array(INT_APP_WORK_COL_DATE_OFFSET + 1)
'    End If
'
'    Set retrieve_dates = New Collection
'
'    Do While rg_app_work.Value <> ""
'        retrieve_dates.add rg_app_work.Value
'        Set rg_app_work = rg_app_work.Offset(1)
'    Loop
'End Function



