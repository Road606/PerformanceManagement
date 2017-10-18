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

Public str_process_step_separator As String

'Public col_kpi_intervals As Collection
Public col_result_pallets As Collection
'Public col_kpi_pallets As Collection '

Public obj_md_dashboard As MDDashboard1
Public obj_mdl_kpi_on_time_result As MDLKPIOnTimeResult

Public Function add_mdl_kpi_on_time_result(obj_mdl As MDLKPIOnTimeResult)
    Set obj_mdl_kpi_on_time_result = obj_mdl
    obj_mdl.add_listener Me
End Function

' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_loading_data_is_about_to_start(obj_date_start As Date, obj_date_end As Date)
    obj_mdl_kpi_on_time_result.load_data_for_interval obj_date_start
End Function

' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_process_record(obj_kpi_on_time_pallet As KPIOnTimePallet)
    Dim obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet
    Dim str_pallet_id As String
    Dim bool_action_create As Boolean

    If obj_kpi_on_time_pallet.str_id = "357020105500768128" Then
        DoEvents
    End If

    bool_action_create = False

    str_pallet_id = retrieve_col_result_pallet_id( _
        obj_kpi_on_time_pallet.str_id, _
        obj_kpi_on_time_pallet.str_place_from, _
        obj_kpi_on_time_pallet.str_place_to)
    On Error GoTo INFO_NEW_PALLET
    Set obj_kpi_on_time_result_pallet = col_result_pallets(str_pallet_id)
    On Error GoTo 0
    
    If Not bool_action_create Then
        col_result_pallets.Remove str_pallet_id
    End If
    update_pallet obj_kpi_on_time_pallet, obj_kpi_on_time_result_pallet

    obj_mdl_kpi_on_time_result.save_record_dynamic obj_kpi_on_time_result_pallet
    Exit Function
INFO_NEW_PALLET:
    Set obj_kpi_on_time_result_pallet = create_pallet(obj_kpi_on_time_pallet)
    bool_action_create = True
    Resume Next
End Function

' # interface method for listening MDLKPIOnTImePallet
Public Function kpi_pallet_loading_data_has_finished(obj_period As Period)
    Dim obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet
    
    For Each obj_kpi_on_time_result_pallet In col_result_pallets
        obj_mdl_kpi_on_time_result.save_record_dynamic obj_kpi_on_time_result_pallet
    Next

    ' clean up
    Set col_result_pallets = New Collection
    obj_mdl_kpi_on_time_result.obj_multi_data_provider.close_providers
End Function

' # interface method for listening MDLKPIOnTimeResult
Public Function kpi_result_process_record(obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet)
    Dim arr_process_step As Variant
    
    arr_process_step = Split(obj_kpi_on_time_result_pallet.str_process_step, str_process_step_separator)
    col_result_pallets.add obj_kpi_on_time_result_pallet, _
        retrieve_col_result_pallet_id( _
            obj_kpi_on_time_result_pallet.str_id, CStr(arr_process_step(0)), CStr(arr_process_step(1)))
End Function

Private Function create_pallet(obj_kpi_on_time_pallet As KPIOnTimePallet) As KPIOnTimeResultPallet
    Dim obj_kpi_on_time_interval As KPIOnTimeInterval

    Set create_pallet = New KPIOnTimeResultPallet
    create_pallet.str_id = obj_kpi_on_time_pallet.str_id
    Set obj_kpi_on_time_interval = obj_kpi_on_time_pallet.obj_kpi_interval
    create_pallet.str_date = DateValue(obj_kpi_on_time_interval.str_id)
    create_pallet.str_time = TimeValue(obj_kpi_on_time_interval.str_id)
    create_pallet.str_shift = obj_kpi_on_time_interval.str_shift
    create_pallet.str_shift_type = obj_kpi_on_time_interval.str_shift_type

    create_pallet.str_building = obj_kpi_on_time_pallet.str_place_building
    create_pallet.str_process_name = obj_kpi_on_time_pallet.str_process_version_name ' # implement conversion to kpi process name
    create_pallet.str_process_step = obj_kpi_on_time_pallet.str_place_from & str_process_step_separator & obj_kpi_on_time_pallet.str_place_to

    create_pallet.str_kpi_id = obj_kpi.str_id
    create_pallet.str_kpi_name = obj_kpi.str_name
End Function

Private Function update_pallet(obj_kpi_on_time_pallet As KPIOnTimePallet, obj_kpi_on_time_result_pallet As KPIOnTimeResultPallet)
    obj_kpi_on_time_result_pallet.str_kpi_status = obj_kpi_on_time_pallet.bool_status_ok
End Function

Private Function retrieve_col_result_pallet_id(str_pallet As String, str_place_from As String, str_place_to As String) As String
    retrieve_col_result_pallet_id = str_pallet & str_place_from & str_place_to
End Function

Private Sub Class_Initialize()
    'Set col_kpi_intervals = New Collection
    Set col_result_pallets = New Collection
    'Set col_kpi_pallets = New Collection
    
    str_process_step_separator = ">"
    
    'Set obj_checkpoint_type_resolver = New CheckpointTypeResolver ' # implement - cross check if this is the correct place where to initialize it
End Sub


