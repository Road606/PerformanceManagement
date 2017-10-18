VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLKPIOnTimePallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLKPIOnTimePallet"

Public STR_DB_TYPE_FILE_EXCEL As String

Public INT_OFFSET_PALLET As Integer
Public INT_OFFSET_PROCESS As Integer
Public INT_OFFSET_PROCESS_VERSION As Integer
Public INT_OFFSET_PROCESS_VERSION_NAME As Integer
Public INT_OFFSET_PROCESS_STATUS As Integer
Public INT_OFFSET_MATERIAL_VENDOR As Integer
Public INT_OFFSET_MATERIAL_BIN_STORAGE_GROUP As Integer
Public INT_OFFSET_PLACE_BUILDING As Integer
Public INT_OFFSET_PLACE_FROM As Integer
Public INT_OFFSET_PLACE_TO As Integer
Public INT_OFFSET_DATE_START As Integer
Public INT_OFFSET_DATE_END As Integer
Public INT_OFFSET_DATE_LAST_CHECKPOINT As Integer
Public INT_OFFSET_INTERVAL_ID As Integer
Public INT_OFFSET_INTERVAL_START As Integer
Public INT_OFFSET_INTERVAL_END As Integer
Public INT_OFFSET_INTERVAL_SHIFT As Integer
Public INT_OFFSET_INTERVAL_SHIFT_TYPE As Integer
Public INT_OFFSET_CHECKPOINT_LIMIT As Integer

'Public INT_APP_WORK_COL_DATE_OFFSET As Integer
'Public STR_APP_WORK_FIRST_RG As String
'
'Public BYTE_STEP_STATUS As Byte

Public obj_multi_data_provider As MultiFileExcelDataProvider
Public obj_current_data_provider_info As FileExcelDataProviderInfo
Public str_static_data_provider_id As String
    ' listeners must implement:
      ' process_record obj_db_data_process: DBDataProcess
Public col_listeners As Collection
    ' listeners must implement:
      ' loading_data_is_about_to_start obj_provider_info: FileExcelDataProviderInfo
'Public col_provider_info_listeners As Collection

Private Sub Class_Initialize()
    STR_DB_TYPE_FILE_EXCEL = "file:excel"

    INT_OFFSET_PALLET = 0
    INT_OFFSET_PROCESS = 1
    INT_OFFSET_PROCESS_VERSION = 2
    INT_OFFSET_PROCESS_VERSION_NAME = 3
    'INT_OFFSET_PROCESS_STATUS = 4
    INT_OFFSET_MATERIAL_VENDOR = 4
    INT_OFFSET_MATERIAL_BIN_STORAGE_GROUP = 5
    INT_OFFSET_PLACE_BUILDING = 6
    INT_OFFSET_PLACE_FROM = 7
    INT_OFFSET_PLACE_TO = 8
    INT_OFFSET_DATE_START = 9
    INT_OFFSET_DATE_END = 10
    INT_OFFSET_DATE_LAST_CHECKPOINT = 11
    INT_OFFSET_INTERVAL_ID = 12
    INT_OFFSET_INTERVAL_START = 13
    INT_OFFSET_INTERVAL_END = 14
    INT_OFFSET_INTERVAL_SHIFT = 15
    INT_OFFSET_INTERVAL_SHIFT_TYPE = 16
    INT_OFFSET_CHECKPOINT_LIMIT = 17
    
    ' app help sheet
'    INT_APP_WORK_COL_DATE_OFFSET = 0
'    STR_APP_WORK_FIRST_RG = "A2"
    
    'Set obj_multi_data_provider = New MultiFileExcelDataProvider
    str_static_data_provider_id = "static"
    Set col_listeners = New Collection
'    Set col_provider_info_listeners = New Collection
End Sub

Public Function set_clear_data()
    obj_multi_data_provider.set_clear_data_before_close_status True
End Function

Public Function reset_clear_data()
    obj_multi_data_provider.set_clear_data_before_close_status False
End Function

Public Function load()
    obj_multi_data_provider.add_listener Me
    obj_multi_data_provider.load_data
End Function

Public Function load_data_for_interval(obj_kpi_on_time_interval As KPIOnTimeInterval)
    Dim obj_data_provider As FileExcelDataProvider
    Dim obj_data_provider_util As Object
    Dim str_provider_id As String
    
    Set obj_data_provider_util = obj_multi_data_provider.obj_data_provider_util
    obj_data_provider_util.str_datetime = obj_kpi_on_time_interval.str_date_time
    str_provider_id = obj_data_provider_util.str_provider_id
    On Error GoTo INFO_PROVIDER_NOT_FOUND
    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(str_provider_id)
    On Error GoTo 0
    obj_data_provider.bool_clear_before_close_status = True
    obj_data_provider.load_data
    obj_data_provider.bool_clear_before_close_status = False
    Exit Function
INFO_PROVIDER_NOT_FOUND:
'    Set obj_data_provider = _
'        obj_multi_data_provider.create_data_provider_save_mode( _
'            str_provider_id, _
'            obj_data_provider_util.retrieve_creation_file_name, _
'            obj_data_provider_util.retrieve_creation_tmpl_file_name)
    'Resume Next
End Function

' # interface method for listening FileExcelDataProvider
Public Function load_record(rg_record As Range)
    Dim obj_kpi_on_time_pallet As KPIOnTimePallet
    Dim obj_kpi_interval As KPIOnTimeInterval
    Dim obj_interval As Interval
    Dim obj_checkpoint As KPIOnTimeLimitCheckpoint
    Dim obj_listener As Object

    ' fill data in a row
    Set obj_kpi_on_time_pallet = New KPIOnTimePallet
      ' pallet level
    obj_kpi_on_time_pallet.str_id = rg_record.Offset(0, INT_OFFSET_PALLET).Value
      ' process
    obj_kpi_on_time_pallet.str_process_id = rg_record.Offset(0, INT_OFFSET_PROCESS).Value
    obj_kpi_on_time_pallet.str_process_version_id = rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION).Value
    obj_kpi_on_time_pallet.str_process_version_name = rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION_NAME).Value
      ' material
    obj_kpi_on_time_pallet.str_material_vendor = rg_record.Offset(0, INT_OFFSET_MATERIAL_VENDOR).Value
    obj_kpi_on_time_pallet.str_material_bin_storage_group = rg_record.Offset(0, INT_OFFSET_MATERIAL_BIN_STORAGE_GROUP).Value
      ' place
    obj_kpi_on_time_pallet.str_place_building = rg_record.Offset(0, INT_OFFSET_PLACE_BUILDING).Value
    obj_kpi_on_time_pallet.str_place_from = rg_record.Offset(0, INT_OFFSET_PLACE_FROM).Value
    obj_kpi_on_time_pallet.str_place_to = rg_record.Offset(0, INT_OFFSET_PLACE_TO).Value
      ' date and time
    obj_kpi_on_time_pallet.str_date_start = rg_record.Offset(0, INT_OFFSET_DATE_START).Value
    obj_kpi_on_time_pallet.str_date_end = rg_record.Offset(0, INT_OFFSET_DATE_END).Value
    obj_kpi_on_time_pallet.str_date_last_checkpoint = rg_record.Offset(0, INT_OFFSET_DATE_LAST_CHECKPOINT).Value
      ' interval
        ' start and end
    Set obj_interval = New Interval
    obj_interval.obj_start = rg_record.Offset(0, INT_OFFSET_INTERVAL_START).Value
    obj_interval.obj_end = rg_record.Offset(0, INT_OFFSET_INTERVAL_END).Value
    Set obj_kpi_interval = New KPIOnTimeInterval
    obj_kpi_interval.obj_interval = obj_interval
    obj_kpi_interval.str_date_time = rg_record.Offset(0, INT_OFFSET_INTERVAL_ID).Value
        ' shift
    obj_kpi_interval.str_shift = rg_record.Offset(0, INT_OFFSET_INTERVAL_SHIFT).Value
    obj_kpi_interval.str_shift_type = rg_record.Offset(0, INT_OFFSET_INTERVAL_SHIFT_TYPE).Value
    Set obj_kpi_on_time_pallet.obj_kpi_interval = obj_kpi_interval
    
      ' checkpoint
    Set obj_checkpoint = New KPIOnTimeLimitCheckpoint
    obj_checkpoint.str_id = rg_record.Offset(0, INT_OFFSET_PLACE_TO).Value
    obj_checkpoint.obj_limit = rg_record.Offset(0, INT_OFFSET_CHECKPOINT_LIMIT).Value
    Set obj_kpi_on_time_pallet.obj_checkpoint = obj_checkpoint
            
    For Each obj_listener In col_listeners
        obj_listener.kpi_pallet_process_record obj_kpi_on_time_pallet
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    Set obj_current_data_provider_info = obj_provider_info
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_pallet_loading_data_is_about_to_start obj_provider_info.obj_period
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_pallet_loading_data_has_finished obj_provider_info.obj_period
    Next
    
    Set obj_current_data_provider_info = Nothing
End Function

Public Function save_record_dynamic(obj_kpi_on_time_pallet As KPIOnTimePallet)
    Dim obj_data_provider As FileExcelDataProvider
    Dim obj_data_provider_util As Object
    'Dim str_provider_id As String

    Set obj_data_provider_util = obj_multi_data_provider.obj_data_provider_util
    obj_data_provider_util.str_datetime = obj_kpi_on_time_pallet.obj_kpi_interval.str_date_time 'obj_interval.obj_start 'obj_kpi_on_time_pallet.str_date_start '& " " & obj_kpi_on_time_pallet.str_time
    'str_provider_id = obj_data_provider_util.str_provider_id
    On Error GoTo INFO_PROVIDER_NOT_FOUND
    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(obj_data_provider_util.str_provider_id)
    On Error GoTo 0

    save_record obj_kpi_on_time_pallet, obj_data_provider.next_data
    Exit Function
INFO_PROVIDER_NOT_FOUND:
    obj_data_provider_util.str_provider_status = retrieve_data_provider_completeness_status(obj_kpi_on_time_pallet.obj_kpi_interval)

    Set obj_data_provider = _
        obj_multi_data_provider.create_data_provider_save_mode( _
            obj_data_provider_util.str_provider_id, _
            obj_data_provider_util.retrieve_creation_file_name, _
            obj_data_provider_util.retrieve_creation_tmpl_file_name, _
            new_const_excel_data_provider.STR_SOURCE_TYPE_EXTERNAL, _
            new_const_excel_data_provider.STR_DATA_PROVIDER_INFO_DYNAMIC)
    'obj_data_provider.obj_info.str_completeness_status = obj_data_provider_util.str_provider_status
    ' open file to start saving data
    obj_data_provider.open_data
    Resume Next
End Function

Public Function save_record_static(obj_data_process_record As DBDataProcess)
    save_record obj_data_process_record, obj_multi_data_provider.retrieve_provider(str_static_data_provider_id).next_data
End Function

Public Function save_record(obj_kpi_on_time_pallet As KPIOnTimePallet, rg_record As Range)
    ' pallet level
    rg_record.Offset(0, INT_OFFSET_PALLET).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_PALLET).Value = obj_kpi_on_time_pallet.str_id
     ' process
    rg_record.Offset(0, INT_OFFSET_PROCESS).Value = obj_kpi_on_time_pallet.str_process_id
    rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION).Value = obj_kpi_on_time_pallet.str_process_version_id
    rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION_NAME).Value = obj_kpi_on_time_pallet.str_process_version_name
      ' material
    rg_record.Offset(0, INT_OFFSET_MATERIAL_VENDOR).Value = obj_kpi_on_time_pallet.str_material_vendor
    rg_record.Offset(0, INT_OFFSET_MATERIAL_BIN_STORAGE_GROUP).Value = obj_kpi_on_time_pallet.str_material_bin_storage_group
      ' place
    rg_record.Offset(0, INT_OFFSET_PLACE_BUILDING).Value = obj_kpi_on_time_pallet.str_place_building
    rg_record.Offset(0, INT_OFFSET_PLACE_FROM).Value = obj_kpi_on_time_pallet.str_place_from
    rg_record.Offset(0, INT_OFFSET_PLACE_TO).Value = obj_kpi_on_time_pallet.str_place_to
      ' date and time
    rg_record.Offset(0, INT_OFFSET_DATE_START).Value = obj_kpi_on_time_pallet.str_date_start
    rg_record.Offset(0, INT_OFFSET_DATE_END).Value = obj_kpi_on_time_pallet.str_date_end
    rg_record.Offset(0, INT_OFFSET_DATE_LAST_CHECKPOINT).Value = obj_kpi_on_time_pallet.str_date_last_checkpoint
      ' interval
        ' start, end
    rg_record.Offset(0, INT_OFFSET_INTERVAL_ID).Value = obj_kpi_on_time_pallet.obj_kpi_interval.str_id
    rg_record.Offset(0, INT_OFFSET_INTERVAL_START).Value = obj_kpi_on_time_pallet.obj_kpi_interval.obj_interval.obj_start
    rg_record.Offset(0, INT_OFFSET_INTERVAL_END).Value = obj_kpi_on_time_pallet.obj_kpi_interval.obj_interval.obj_end
        ' shift
    rg_record.Offset(0, INT_OFFSET_INTERVAL_SHIFT).Value = obj_kpi_on_time_pallet.obj_kpi_interval.str_shift
    rg_record.Offset(0, INT_OFFSET_INTERVAL_SHIFT_TYPE).Value = obj_kpi_on_time_pallet.obj_kpi_interval.str_shift_type
      ' checkpoint
    rg_record.Offset(0, INT_OFFSET_CHECKPOINT_LIMIT).NumberFormat = "[$-F400]h:mm:ss AM/PM"
    rg_record.Offset(0, INT_OFFSET_CHECKPOINT_LIMIT).Value = obj_kpi_on_time_pallet.obj_checkpoint.obj_limit
End Function

Public Function close_data()
    obj_multi_data_provider.close_providers
End Function

Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function

Private Function retrieve_data_provider_completeness_status(obj_kpi_interval As KPIOnTimeInterval) As String
    If obj_kpi_interval.bool_is_finished Then
        retrieve_data_provider_completeness_status = new_const_excel_data_provider.STR_DATA_PROVIDER_INFO_STATUS_CLOSED
    Else
        retrieve_data_provider_completeness_status = new_const_excel_data_provider.STR_DATA_PROVIDER_INFO_STATUS_OPENED
    End If
End Function
