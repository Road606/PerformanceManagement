VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLKPIOnTimeResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLKPIOnTimeResult"

Public STR_DB_TYPE_FILE_EXCEL As String

Public INT_OFFSET_PALLET As Integer
Public INT_OFFSET_DATE As Integer
Public INT_OFFSET_TIME As Integer
Public INT_OFFSET_SHIFT As Integer
Public INT_OFFSET_SHIFT_TYPE As Integer
Public INT_OFFSET_BUILDING As Integer
Public INT_OFFSET_PROCESS_NAME As Integer
Public INT_OFFSET_PROCESS_STEP As Integer
Public INT_OFFSET_KPI_ID As Integer
Public INT_OFFSET_KPI_NAME As Integer
Public INT_OFFSET_KPI_STATUS As Integer

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
    INT_OFFSET_DATE = 1
    INT_OFFSET_TIME = 2
    INT_OFFSET_SHIFT = 3
    INT_OFFSET_SHIFT_TYPE = 4
    INT_OFFSET_BUILDING = 5
    INT_OFFSET_PROCESS_NAME = 6
    INT_OFFSET_PROCESS_STEP = 7
    INT_OFFSET_KPI_ID = 8
    INT_OFFSET_KPI_NAME = 9
    INT_OFFSET_KPI_STATUS = 10
        
    str_static_data_provider_id = "static"
    Set col_listeners = New Collection
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

Public Function load_data_for_interval(obj_date_start As Date)
    Dim obj_data_provider As FileExcelDataProvider
    Dim obj_data_provider_util As FileExcelDataProviderUtil
    Dim str_provider_id As String
    
    Set obj_data_provider_util = obj_multi_data_provider.obj_data_provider_util
    obj_data_provider_util.str_datetime = obj_date_start
    str_provider_id = obj_data_provider_util.str_provider_id
    On Error GoTo INFO_PROVIDER_NOT_FOUND
    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(str_provider_id)
    On Error GoTo 0
    
    obj_data_provider.bool_clear_before_close_status = True
    On Error GoTo INFO_FILE_NOT_EXISTS
    obj_data_provider.load_data
    On Error GoTo 0
    obj_multi_data_provider.remove_provider obj_data_provider
    Exit Function
INFO_PROVIDER_NOT_FOUND:
    Set obj_data_provider = _
        obj_multi_data_provider.create_data_provider_read_mode( _
            str_provider_id, obj_data_provider_util.retrieve_creation_file_name)
    obj_data_provider.bool_save_mode_on = True ' to be able save the file as empty one when loading is finished
    Resume Next
INFO_FILE_NOT_EXISTS:
    ' file not exists yet => remove provider
    obj_multi_data_provider.remove_provider obj_data_provider
End Function

' # interface method for listening FileExcelDataProvider
Public Function load_record(rg_record As Range)
    Dim obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet
    Dim obj_listener As Object

    ' fill data in a row
    Set obj_kpi_on_time_result_pallet = New KPIOnTimeResultPallet
      ' pallet level
    obj_kpi_on_time_result_pallet.str_id = rg_record.Offset(0, INT_OFFSET_PALLET).Value
      ' date and time
    obj_kpi_on_time_result_pallet.str_date = rg_record.Offset(0, INT_OFFSET_DATE).Value
    obj_kpi_on_time_result_pallet.str_time = rg_record.Offset(0, INT_OFFSET_TIME).Value
      ' shift
    obj_kpi_on_time_result_pallet.str_shift = rg_record.Offset(0, INT_OFFSET_SHIFT).Value
    obj_kpi_on_time_result_pallet.str_shift_type = rg_record.Offset(0, INT_OFFSET_SHIFT_TYPE).Value
      ' place
    obj_kpi_on_time_result_pallet.str_building = rg_record.Offset(0, INT_OFFSET_BUILDING).Value
      ' process
    obj_kpi_on_time_result_pallet.str_process_name = rg_record.Offset(0, INT_OFFSET_PROCESS_NAME).Value
    obj_kpi_on_time_result_pallet.str_process_step = rg_record.Offset(0, INT_OFFSET_PROCESS_STEP).Value
      ' kpi
    obj_kpi_on_time_result_pallet.str_kpi_id = rg_record.Offset(0, INT_OFFSET_KPI_ID).Value
    obj_kpi_on_time_result_pallet.str_kpi_name = rg_record.Offset(0, INT_OFFSET_KPI_NAME).Value
    obj_kpi_on_time_result_pallet.str_kpi_status = rg_record.Offset(0, INT_OFFSET_KPI_STATUS).Value
    
    For Each obj_listener In col_listeners
        obj_listener.kpi_result_process_record obj_kpi_on_time_result_pallet
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    Set obj_current_data_provider_info = obj_provider_info
    
    For Each obj_listener In col_listeners
        obj_listener.loading_data_is_about_to_start obj_provider_info
    Next
End Function

' # interface method for listening MultiFileExcelDataProvider
Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object

    Set obj_current_data_provider_info = obj_provider_info
    
    For Each obj_listener In col_listeners
        obj_listener.loading_data_is_about_to_start obj_provider_info.obj_period
    Next
End Function

Public Function save_record_dynamic(obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet)
    Dim obj_data_provider As FileExcelDataProvider
    Dim obj_data_provider_util As FileExcelDataProviderUtil
    'Dim str_provider_id As String

    Set obj_data_provider_util = obj_multi_data_provider.obj_data_provider_util
    obj_data_provider_util.str_datetime = obj_kpi_on_time_result_pallet.str_date & " " & obj_kpi_on_time_result_pallet.str_time
    'str_provider_id = obj_data_provider_util.str_provider_id
    On Error GoTo INFO_PROVIDER_NOT_FOUND
    Set obj_data_provider = obj_multi_data_provider.retrieve_provider(obj_data_provider_util.str_provider_id)
    On Error GoTo 0

    save_record obj_kpi_on_time_result_pallet, obj_data_provider.next_data
    Exit Function
INFO_PROVIDER_NOT_FOUND:
    Set obj_data_provider = _
        obj_multi_data_provider.create_data_provider_save_mode( _
            obj_data_provider_util.str_provider_id, _
            obj_data_provider_util.retrieve_creation_file_name, _
            obj_data_provider_util.retrieve_creation_tmpl_file_name)
    ' open file to start saving data
    obj_data_provider.open_data
    Resume Next
End Function

Public Function save_record_static(obj_data_process_record As DBDataProcess)
    save_record obj_data_process_record, obj_multi_data_provider.retrieve_provider(str_static_data_provider_id).next_data
End Function

Public Function save_record(obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet, rg_record As Range)
    ' pallet level
    rg_record.Offset(0, INT_OFFSET_PALLET).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_PALLET).Value = obj_kpi_on_time_result_pallet.str_id
    ' date and time
    rg_record.Offset(0, INT_OFFSET_DATE).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_DATE).Value = obj_kpi_on_time_result_pallet.str_date
    rg_record.Offset(0, INT_OFFSET_TIME).NumberFormat = "@"
    rg_record.Offset(0, INT_OFFSET_TIME).Value = obj_kpi_on_time_result_pallet.str_time
    ' shift
    rg_record.Offset(0, INT_OFFSET_SHIFT).Value = obj_kpi_on_time_result_pallet.str_shift
    rg_record.Offset(0, INT_OFFSET_SHIFT_TYPE).Value = obj_kpi_on_time_result_pallet.str_shift_type
    ' place
    rg_record.Offset(0, INT_OFFSET_BUILDING).Value = obj_kpi_on_time_result_pallet.str_building
    ' process
    rg_record.Offset(0, INT_OFFSET_PROCESS_NAME).Value = obj_kpi_on_time_result_pallet.str_process_name
    rg_record.Offset(0, INT_OFFSET_PROCESS_STEP).Value = obj_kpi_on_time_result_pallet.str_process_step
    ' kpi
    rg_record.Offset(0, INT_OFFSET_KPI_ID).Value = obj_kpi_on_time_result_pallet.str_kpi_id
    rg_record.Offset(0, INT_OFFSET_KPI_NAME).Value = obj_kpi_on_time_result_pallet.str_kpi_name
    rg_record.Offset(0, INT_OFFSET_KPI_STATUS).Value = obj_kpi_on_time_result_pallet.str_kpi_status
End Function

Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function
