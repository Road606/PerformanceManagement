VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtrlKPIOnTimeResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STR_CLASS_NAME As String = "CtrlKPIOnTimeResult" ' # implement should be changed to KPIOnTimeResult

Public str_id As String
Public obj_kpi As Object

'Public col_kpi_intervals As Collection
Public col_result_pallets As Collection
'Public col_kpi_pallets As Collection '

Public obj_md_dashboard As MDDashboard1
Public obj_mdl_kpi_on_time_result As MDLKPIOnTimeResult

'Public obj_checkpoint_type_resolver As CheckpointTypeResolver

Public Function kpi_pallet_loading_data_is_about_to_start(obj_date_start As Date, obj_date_end As Date)
    obj_mdl_kpi_on_time_result.load_data_for_interval obj_date_start
End Function

Public Function kpi_pallet_process_record(obj_kpi_on_time_pallet As KPIOnTimePallet)
    
    Dim obj_kpi_on_time_interval As KPIOnTimeInterval
    Dim obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet

    Set obj_kpi_on_time_result_pallet = New KPIOnTimeResultPallet
    obj_kpi_on_time_result_pallet.str_id = obj_kpi_on_time_pallet.str_id
    Set obj_kpi_on_time_interval = obj_kpi_on_time_pallet.obj_kpi_interval
    obj_kpi_on_time_result_pallet.str_date = DateValue(obj_kpi_on_time_interval.str_id)
    obj_kpi_on_time_result_pallet.str_time = TimeValue(obj_kpi_on_time_interval.str_id)
    obj_kpi_on_time_result_pallet.str_shift = obj_kpi_on_time_interval.str_shift
    obj_kpi_on_time_result_pallet.str_shift_type = obj_kpi_on_time_interval.str_shift_type

    obj_kpi_on_time_result_pallet.str_building = obj_kpi_on_time_pallet.str_place_building
    obj_kpi_on_time_result_pallet.str_process_name = obj_kpi_on_time_pallet.str_process_version_name ' # implement conversion to kpi process name
    obj_kpi_on_time_result_pallet.str_process_step = obj_kpi_on_time_pallet.str_place_from & ">" & obj_kpi_on_time_pallet.str_place_to

    obj_kpi_on_time_result_pallet.str_kpi_id = obj_kpi.str_id
    obj_kpi_on_time_result_pallet.str_kpi_name = obj_kpi.str_name
    obj_kpi_on_time_result_pallet.str_kpi_status = obj_kpi_on_time_pallet.bool_status_ok

    obj_mdl_kpi_on_time_result.save_record_dynamic obj_kpi_on_time_result_pallet
End Function

' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_loading_data_has_finished(obj_period As Period)
    ' clean up
    Set col_result_pallets = New Collection
    obj_mdl_kpi_on_time_result.obj_multi_data_provider.close_providers
End Function

Public Function process_result_record(obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet)
    col_result_pallets.add obj_kpi_on_time_result_pallet, obj_kpi_on_time_result_pallet.str_id
End Function




Private Sub Class_Initialize()
    'Set col_kpi_intervals = New Collection
    Set col_result_pallets = New Collection
    'Set col_kpi_pallets = New Collection
    
    'Set obj_checkpoint_type_resolver = New CheckpointTypeResolver ' # implement - cross check if this is the correct place where to initialize it
End Sub


