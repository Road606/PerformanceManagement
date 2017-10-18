VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtrlKPIOnTimePallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STR_CLASS_NAME As String = "CtrlKPIOnTimePallet" ' # implement should be changed to KPIOnTimeResult

Public str_id As String
Public obj_kpi As Object

' data
Public obj_current_provider_info As FileExcelDataProviderInfo

Public col_db_data_process_unfinished As Collection
Public col_kpi_intervals As Collection
Public col_kpi_pallets As Collection '

Public bool_db_data_process_unfinished_loaded As Boolean

' master data
Public obj_md_dashboard As MDDashboard1
Public obj_mdl_kpi_on_time_pallet As MDLKPIOnTimePallet

Public obj_checkpoint_type_resolver As CheckpointTypeResolver

' # method is called from MDLProcessDashboardListener
Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_db_data_process As DBDataProcess

    Set obj_current_provider_info = obj_provider_info

    If obj_provider_info.str_loading_type = new_const_excel_data_provider.STR_DATA_PROVIDER_INFO_DYNAMIC Then
        create_intervals obj_provider_info.obj_period.str_start, obj_provider_info.obj_period.str_end
        
        'If Not bool_db_data_process_unfinished_loaded Then
'            For Each obj_db_data_process In col_db_data_process_unfinished
'                process_db_data_process_record obj_db_data_process
'            Next
'
'            bool_db_data_process_unfinished_loaded = True
        'End If
    Else
        Set col_db_data_process_unfinished = New Collection
    End If
End Function

' # method is called from MDLProcessDashboardListener
Public Function process_record(obj_db_data_process As DBDataProcess)
    If obj_current_provider_info.str_loading_type = new_const_excel_data_provider.STR_DATA_PROVIDER_INFO_DYNAMIC Then
        process_db_data_process_record obj_db_data_process
    Else
        process_db_data_process_unfinished obj_db_data_process
    End If
End Function

' # method is called from MDLProcessDashboardListener
Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    If obj_provider_info.str_loading_type = new_const_excel_data_provider.STR_DATA_PROVIDER_INFO_DYNAMIC Then
        save obj_provider_info
        
        ' clean up
        Set col_kpi_intervals = New Collection
        obj_mdl_kpi_on_time_pallet.close_data
    End If

    Set obj_current_provider_info = Nothing
End Function

Public Function create_intervals(obj_period_start As Date, obj_period_end As Date)
    Dim obj_interval As Interval ' standard interval
    Dim obj_process_interval As ProcessInterval ' master data for process
    Dim obj_kpi_on_time_interval As KPIOnTimeInterval ' specific interval for dashboard usage
    Dim obj_kpi_on_time_interval_last As KPIOnTimeInterval ' specific interval for dashboard usage - last created
    Dim col_intervals As Collection
    
    For Each obj_process_interval In obj_kpi.obj_process.col_intervals
        Set col_intervals = fact_interval.create_intervals_hourly( _
            obj_period_start, obj_period_end, _
            obj_process_interval.int_lower_limit, obj_process_interval.int_upper_limit, obj_process_interval.int_hour_offset)
            
        For Each obj_interval In col_intervals
            Set obj_kpi_on_time_interval = New KPIOnTimeInterval
            obj_kpi_on_time_interval.obj_interval = obj_interval
            obj_kpi_on_time_interval.str_date_time = DateValue(obj_interval.obj_start) + _
                TimeSerial(Hour(obj_interval.obj_start) + obj_process_interval.int_hour_offset, obj_process_interval.int_minute, 0)
            obj_kpi_on_time_interval.bool_is_finished = CBool(obj_process_interval.int_is_complete)
            obj_kpi_on_time_interval.str_shift = hndl_master_shift.find_shift(Format(obj_interval.obj_start, "DD.MM.YYYY-h")) ' # implement through new class
            obj_kpi_on_time_interval.str_shift_type = hndl_master_shift.find_type(Hour(obj_interval.obj_start))
            col_kpi_intervals.add obj_kpi_on_time_interval, obj_kpi_on_time_interval.str_id
            
            ' create connection between kpi intervals
            If Not obj_kpi_on_time_interval_last Is Nothing Then
                Set obj_kpi_on_time_interval.obj_previous_interval = obj_kpi_on_time_interval_last
                Set obj_kpi_on_time_interval_last.obj_next_interval = obj_kpi_on_time_interval
            End If
            
            Set obj_kpi_on_time_interval_last = obj_kpi_on_time_interval
        Next
    Next
End Function

Public Function process_db_data_process_unfinished(obj_data_process As DBDataProcess)
       col_db_data_process_unfinished.add obj_data_process, obj_data_process.str_pallet
End Function

'' # method is called from MDLProcessDashboardListener
Public Function process_db_data_process_record(obj_data_process As DBDataProcess)
    Dim obj_kpi_on_time_interval As KPIOnTimeInterval
    Dim obj_kpi_on_time_next_interval As KPIOnTimeInterval
    Dim obj_kpi_on_time_pallet As KPIOnTimePallet
    Dim obj_kpi_on_time_pallet_step As KPIOnTimePallet
    Dim obj_kpi_on_time_pallet_steps As Collection
    Dim obj_kpi_on_time_pallet_step_interval As KPIOnTimePallet
    Dim obj_kpi_on_time_pallet_full_live_interval As KPIOnTimePallet
    Dim str_date_end_for_status_calculation As String

    If obj_data_process.str_pallet = "357020105501109173" Then
        DoEvents
    End If

    ' check if kpi on time is interested into current obj_data_process
    If obj_kpi.obj_process.str_process_id = obj_data_process.str_creation_id Then
        On Error GoTo INFO_NEW_PALLET
        Set obj_kpi_on_time_pallet = col_kpi_pallets(obj_data_process.str_pallet)
        On Error GoTo 0
        
        If Not obj_kpi_on_time_pallet.obj_checkpoint Is Nothing Then ' #implement - don't know if this IF statement is necesarry
            Set obj_kpi_on_time_pallet_steps = update_kpi_on_time_pallet(obj_kpi_on_time_pallet, obj_data_process)
        
            For Each obj_kpi_on_time_interval In col_kpi_intervals
                For Each obj_kpi_on_time_pallet_step In obj_kpi_on_time_pallet_steps
                    ' find date end for pallet
                    If obj_kpi_on_time_pallet_step.str_date_end = "" Then
                        str_date_end_for_status_calculation = obj_kpi_on_time_interval.obj_interval.obj_end
                    Else
                        str_date_end_for_status_calculation = obj_kpi_on_time_pallet_step.str_date_end
                    End If
                    
                    ' check if pallet belongs to the interval
                    If obj_kpi_on_time_interval.is_valid_interval(obj_kpi_on_time_pallet_step.str_date_last_checkpoint, str_date_end_for_status_calculation) Then
                        On Error GoTo INFO_NEW_PALLET_IN_INTERVAL
                        Set obj_kpi_on_time_pallet_step_interval = obj_kpi_on_time_interval.col_pallets_step(retrieve_pallet_id_interval(obj_kpi_on_time_pallet_step))
                        On Error GoTo 0
                        
                        ' update obj_kpi_on_time_pallet_step_interval
                        update_kpi_on_time_pallet_step_interval _
                            obj_kpi_on_time_interval, obj_kpi_on_time_pallet_step, _
                            obj_kpi_on_time_pallet_step_interval, obj_data_process, _
                            str_date_end_for_status_calculation
                        
                        ' update full live pallet
                        Set obj_kpi_on_time_pallet_full_live_interval = obj_kpi_on_time_pallet_step_interval.create_copy
                        
                        ' # implement - when pallet step interval is closed then remove this specific pallet step from following intervals
                        Set obj_kpi_on_time_next_interval = obj_kpi_on_time_interval.obj_next_interval
                        
                        Do While Not obj_kpi_on_time_next_interval Is Nothing
                            On Error GoTo INFO_TRY_TO_REMOVE
                            obj_kpi_on_time_next_interval.col_pallets_step.Remove retrieve_pallet_id_interval(obj_kpi_on_time_pallet_step_interval)
                            obj_kpi_on_time_next_interval.col_pallets_full_live.Remove obj_kpi_on_time_pallet_full_live_interval.str_id 'retrieve_pallet_id_interval(obj_kpi_on_time_pallet_full_live_interval)
                            On Error GoTo 0
                            Set obj_kpi_on_time_next_interval = obj_kpi_on_time_next_interval.obj_next_interval
                        Loop
                    End If
                Next
                
                ' create
                If Not obj_kpi_on_time_pallet_full_live_interval Is Nothing Then
                    On Error GoTo INFO_NEW_PALLET_IN_FULL_LIVE_INTERVAL
                    obj_kpi_on_time_interval.col_pallets_full_live.add obj_kpi_on_time_pallet_full_live_interval, obj_kpi_on_time_pallet_full_live_interval.str_id 'retrieve_pallet_id_interval(obj_kpi_on_time_pallet_full_live_interval)
                    On Error GoTo 0
                    ' reset
                    Set obj_kpi_on_time_pallet_full_live_interval = Nothing 'obj_kpi_on_time_pallet_full_live_interval.create_copy
                End If
                
            Next
        End If
    End If
    
    Exit Function
INFO_NEW_PALLET:
    Set obj_kpi_on_time_pallet = create_kpi_on_time_pallet(obj_data_process)
    If Not obj_kpi_on_time_pallet Is Nothing Then ' it could happen if checkpoint for current place is not found or set - this behavior is required if only part of process should be measured by KPI
        Resume Next
    End If
    Exit Function
INFO_NEW_PALLET_IN_INTERVAL:
    Set obj_kpi_on_time_pallet_step_interval = obj_kpi_on_time_pallet_step.create_copy
    Set obj_kpi_on_time_pallet_step_interval.obj_kpi_interval = obj_kpi_on_time_interval
    obj_kpi_on_time_interval.col_pallets_step.add obj_kpi_on_time_pallet_step_interval, retrieve_pallet_id_interval(obj_kpi_on_time_pallet_step_interval)
    Resume Next
INFO_NEW_PALLET_IN_FULL_LIVE_INTERVAL:
    obj_kpi_on_time_interval.col_pallets_full_live.Remove obj_kpi_on_time_pallet_full_live_interval.str_id
    obj_kpi_on_time_interval.col_pallets_full_live.add obj_kpi_on_time_pallet_full_live_interval, obj_kpi_on_time_pallet_full_live_interval.str_id
    Resume Next
INFO_TRY_TO_REMOVE:
    Resume Next
End Function

Public Function create_kpi_on_time_pallet(obj_data_process As DBDataProcess) As KPIOnTimePallet
    Set create_kpi_on_time_pallet = New KPIOnTimePallet
    create_kpi_on_time_pallet.str_id = obj_data_process.str_pallet
    create_kpi_on_time_pallet.str_process_id = obj_data_process.str_creation_id ' # implement - creation rename to process id
    create_kpi_on_time_pallet.str_process_version_id = obj_data_process.str_version ' # implement - version rename to process version
    create_kpi_on_time_pallet.str_process_version_name = obj_kpi.col_process_versions(obj_data_process.str_version).str_name ' # implement - version rename to process version
    create_kpi_on_time_pallet.str_material_vendor = obj_data_process.str_material_vendor
    create_kpi_on_time_pallet.str_material_bin_storage_group = obj_data_process.str_material_bin_storage_group
    create_kpi_on_time_pallet.str_building_source = _
        obj_checkpoint_type_resolver.retrieve_building(obj_data_process.str_source, obj_data_process.str_source_type)
    create_kpi_on_time_pallet.str_building_destination = _
        obj_checkpoint_type_resolver.retrieve_building(obj_data_process.str_destination, obj_data_process.str_destination_type)
        
    create_kpi_on_time_pallet.str_place_from = bin_place_grp.get_place_grp(obj_data_process.str_bin_from)
    create_kpi_on_time_pallet.str_date_start = obj_data_process.str_date_start
    create_kpi_on_time_pallet.str_date_last_checkpoint = obj_data_process.str_date_start
    
    ' retrieve limit
    Set create_kpi_on_time_pallet.obj_checkpoint = retrieve_checkpoint(create_kpi_on_time_pallet)
    If create_kpi_on_time_pallet.obj_checkpoint Is Nothing Then
        Set create_kpi_on_time_pallet = Nothing
    Else
        col_kpi_pallets.add create_kpi_on_time_pallet, create_kpi_on_time_pallet.str_id
    End If
    
End Function

Public Function update_kpi_on_time_pallet(obj_kpi_on_time_pallet As KPIOnTimePallet, obj_data_process As DBDataProcess) As Collection
    Dim str_place_to As String
    Dim obj_checkpoint As KPIOnTimeLimitCheckpoint
    Dim obj_kpi_on_time_pallet_new As KPIOnTimePallet
    Dim obj_kpi_on_time_pallet_check As KPIOnTimePallet
    
    Set update_kpi_on_time_pallet = New Collection
    
    obj_kpi_on_time_pallet.byte_process_status = obj_data_process.byte_process_status

    ' check if transaction is complete
    If obj_data_process.str_date_end <> "" Then
        str_place_to = bin_place_grp.get_place_grp(obj_data_process.str_bin_to)
        
        If str_place_to = obj_kpi_on_time_pallet.obj_checkpoint.str_next_checkpoint_id Then
            ' checkpoint is found and came in expected order
            obj_kpi_on_time_pallet.str_date_end = obj_data_process.str_date_end
            obj_kpi_on_time_pallet.str_place_to = str_place_to
            'obj_kpi_on_time_pallet.str_place_building = bin.get_building(obj_data_process.str_bin_to)
            obj_kpi_on_time_pallet.str_place_building = _
                obj_checkpoint_type_resolver.retrieve_building_for_checkpoint( _
                    obj_kpi_on_time_pallet.str_building_source, const_place_type.STR_IMPLICIT, _
                    obj_kpi_on_time_pallet.str_building_destination, const_place_type.STR_IMPLICIT, _
                    obj_kpi_on_time_pallet.obj_checkpoint)
            
            ' create kpi pallet step 1
            update_kpi_on_time_pallet.add obj_kpi_on_time_pallet.create_copy
            ' check if pallet is still on the way to the destination or it already reached its destination
            'If obj_data_process.byte_process_status = new_ctrl_process.BYTE_STATUS_OPEN Then
            If obj_data_process.byte_process_step_status = new_db_process_step.BYTE_PROCESS_STATUS_CONTINUE Then
                Set obj_checkpoint = obj_kpi_on_time_pallet.obj_checkpoint
                update_kpi_on_time_pallet_to_next_checkpoint obj_kpi_on_time_pallet
                
                If Not obj_checkpoint.str_id = obj_kpi_on_time_pallet.obj_checkpoint.str_id Then
                    ' create kpi pallet step 2
                    update_kpi_on_time_pallet.add obj_kpi_on_time_pallet.create_copy
                End If
            End If
        Else
            ' check if checkpoint with new str_place_to exists
            Set obj_kpi_on_time_pallet_check = obj_kpi_on_time_pallet.create_copy
            obj_kpi_on_time_pallet_check.str_place_to = str_place_to
            update_kpi_on_time_pallet_to_next_checkpoint obj_kpi_on_time_pallet_check
            
            ' if checkpoints equals then it means that checkpoint for new str_place_to doesn't exist (update to next checkpoint left pallet unchanged if checkpoint doesn't exist)
            If Not obj_kpi_on_time_pallet.obj_checkpoint.str_id = obj_kpi_on_time_pallet_check.obj_checkpoint.str_id Then
                ' move pallet from current checkpoint to the next closest checkpoint
                Set obj_kpi_on_time_pallet.obj_checkpoint = obj_kpi_on_time_pallet.obj_checkpoint.obj_next_limit
                obj_kpi_on_time_pallet.str_date_end = obj_data_process.str_date_end
                obj_kpi_on_time_pallet.str_place_to = obj_kpi_on_time_pallet.obj_checkpoint.str_id 'str_place_to
                'obj_kpi_on_time_pallet.str_place_building = bin.get_building(obj_data_process.str_bin_to) ' # implement - should be based on checkpoint type (local, remote) and destination
                obj_kpi_on_time_pallet.str_place_building = _
                    obj_checkpoint_type_resolver.retrieve_building_for_checkpoint( _
                        obj_kpi_on_time_pallet.str_building_source, const_place_type.STR_IMPLICIT, _
                        obj_kpi_on_time_pallet.str_building_destination, const_place_type.STR_IMPLICIT, _
                        obj_kpi_on_time_pallet.obj_checkpoint)
                ' create kpi pallet step 1
                update_kpi_on_time_pallet.add obj_kpi_on_time_pallet.create_copy
                
                ' move pallet to new checkpoint
                ' create kpi pallet step 2
                Set obj_kpi_on_time_pallet.obj_checkpoint = obj_kpi_on_time_pallet_check.obj_checkpoint
                obj_kpi_on_time_pallet.str_place_from = obj_kpi_on_time_pallet_check.obj_checkpoint.obj_previous_limit.str_id
                obj_kpi_on_time_pallet.str_place_to = str_place_to
                'obj_kpi_on_time_pallet.str_place_building = bin.get_building(obj_data_process.str_bin_to)
                obj_kpi_on_time_pallet.str_place_building = _
                    obj_checkpoint_type_resolver.retrieve_building_for_checkpoint( _
                        obj_kpi_on_time_pallet.str_building_source, const_place_type.STR_IMPLICIT, _
                        obj_kpi_on_time_pallet.str_building_destination, const_place_type.STR_IMPLICIT, _
                        obj_kpi_on_time_pallet.obj_checkpoint)
                update_kpi_on_time_pallet.add obj_kpi_on_time_pallet.create_copy
                                
                ' check if pallet is still on the way to the destination
                'If obj_data_process.byte_process_status = new_ctrl_process.BYTE_STATUS_OPEN Then
                If obj_data_process.byte_process_step_status = new_db_process_step.BYTE_PROCESS_STATUS_CONTINUE Then
                    Set obj_checkpoint = obj_kpi_on_time_pallet.obj_checkpoint
                    update_kpi_on_time_pallet_to_next_checkpoint obj_kpi_on_time_pallet
                    
                    If Not obj_checkpoint.str_id = obj_kpi_on_time_pallet.obj_checkpoint.str_id Then
                        ' create kpi pallet step 3
                        update_kpi_on_time_pallet.add obj_kpi_on_time_pallet.create_copy
                    End If
                End If
            Else
                update_kpi_on_time_pallet.add obj_kpi_on_time_pallet.create_copy
            End If
        End If
    Else
        update_kpi_on_time_pallet.add obj_kpi_on_time_pallet.create_copy
    End If
End Function

Public Function update_kpi_on_time_pallet_to_next_checkpoint(obj_kpi_on_time_pallet As KPIOnTimePallet)
    Dim obj_kpi_on_time_pallet_original As KPIOnTimePallet
    Dim obj_checkpoint As KPIOnTimeLimitCheckpoint
    
    ' save original version of pallet in case for new str_place_to checkpoint doesn't exist
    Set obj_kpi_on_time_pallet_original = obj_kpi_on_time_pallet.create_copy
    
    ' update first part of pallet step
    obj_kpi_on_time_pallet.str_place_from = obj_kpi_on_time_pallet.str_place_to
    obj_kpi_on_time_pallet.str_date_last_checkpoint = obj_kpi_on_time_pallet.str_date_end
    
    ' reset second part of pallet step
    obj_kpi_on_time_pallet.str_date_end = ""
    obj_kpi_on_time_pallet.str_place_to = ""
    obj_kpi_on_time_pallet.str_place_building = ""
    ' load new checkpoint for new str_place_from
    Set obj_kpi_on_time_pallet.obj_checkpoint = retrieve_checkpoint(obj_kpi_on_time_pallet)
    
    ' if checkpoint for new str_place_to doesn't exist then pallet stays on its current place/checkpoint
    If obj_kpi_on_time_pallet.obj_checkpoint Is Nothing Then
        Set obj_kpi_on_time_pallet = obj_kpi_on_time_pallet_original
    End If
End Function

Public Function update_kpi_on_time_pallet_step_interval( _
        obj_kpi_on_time_interval As KPIOnTimeInterval, _
        obj_kpi_on_time_pallet_step As KPIOnTimePallet, _
        obj_kpi_on_time_pallet_step_interval As KPIOnTimePallet, _
        obj_data_process As DBDataProcess, _
        str_date_end_for_status_calculation)
    ' process status
    obj_kpi_on_time_pallet_step_interval.byte_process_status = obj_kpi_on_time_pallet_step.byte_process_status
    ' place
    ' # implement - if str_place_to = "" then take place_to from next checkpoint
    If obj_kpi_on_time_pallet_step.str_place_to = "" Then
        obj_kpi_on_time_pallet_step_interval.str_place_to = obj_kpi_on_time_pallet_step_interval.obj_checkpoint.str_next_checkpoint_id
        obj_kpi_on_time_pallet_step_interval.str_place_building = _
            obj_checkpoint_type_resolver.retrieve_building_for_checkpoint( _
                obj_kpi_on_time_pallet_step.str_building_source, const_place_type.STR_IMPLICIT, _
                obj_kpi_on_time_pallet_step.str_building_destination, const_place_type.STR_IMPLICIT, _
                obj_kpi_on_time_pallet_step_interval.obj_checkpoint)
        ' # implement - retrieve building from obj_data_process destination property
    Else
        obj_kpi_on_time_pallet_step_interval.str_place_to = obj_kpi_on_time_pallet_step.str_place_to
        obj_kpi_on_time_pallet_step_interval.str_place_building = obj_kpi_on_time_pallet_step.str_place_building
    End If
    ' date end
    If CDate(str_date_end_for_status_calculation) > obj_kpi_on_time_interval.obj_interval.obj_end Then
        obj_kpi_on_time_pallet_step_interval.str_date_end = obj_kpi_on_time_interval.obj_interval.obj_end
    Else
        obj_kpi_on_time_pallet_step_interval.str_date_end = str_date_end_for_status_calculation
    End If
End Function

'Public Function retrieve_checkpoint(obj_kpi_on_time_pallet As KPIOnTimePallet, obj_data_process As DBDataProcess) As KPIOnTimeLimitCheckpoint
Public Function retrieve_checkpoint(obj_kpi_on_time_pallet As KPIOnTimePallet) As KPIOnTimeLimitCheckpoint
    Dim bool_only_default As Boolean ' if during searching one part of limit is not found then every following part is set to default

    'Dim str_process_version_id As String
    Dim str_place_checkpoint As String
    Dim obj_vendor As Vendor
    Dim obj_material_group_wh_resolver As MaterialGroupWHResolver
    Dim obj_condition_process_material As ConditionProcessMaterial
    
    Dim obj_kpi_limit_process As KPIOnTimeLimitProc
    Dim obj_kpi_limit_process_version As KPIOnTimeLimitProcVersion
    Dim obj_kpi_limit_mat_grp_wh As KPIOnTimeLimitMatGrpWH
    Dim obj_kpi_limit_source As KPIOnTimeLimitSource
    
    'Dim obj_mdl_process_version_master As New MDLProcessVersionMaster
    'Dim obj_mdl_material_group_wh As New MDLMaterialGroupWH
    
    ' # implement - all retrievals must be error handled - these will lead to warning error
    Set obj_kpi_limit_process = obj_kpi.col_limits_process(obj_kpi_on_time_pallet.str_process_id)
    Set obj_kpi_limit_process_version = obj_kpi_limit_process.col_limits_version(obj_kpi_on_time_pallet.str_process_version_id)
    
    On Error GoTo WARN_NOT_FOUND_MATERIAL_GROUP_WH_RESOLVER
    Set obj_vendor = obj_md_dashboard.col_material_group_wh_resolvers(obj_kpi_on_time_pallet.str_material_vendor)
    Set obj_material_group_wh_resolver = obj_vendor.retrieve_material_group_wh_resolver(obj_kpi_on_time_pallet.str_material_bin_storage_group)
    Set obj_kpi_limit_mat_grp_wh = obj_kpi_limit_process_version.col_limits_mat_grp_wh(obj_material_group_wh_resolver.str_material_group_wh_id)
    On Error GoTo 0
    
    If Not bool_only_default Then
        On Error GoTo WARN_NOT_FOUND_CONDITION_PROCESS_MATERIAL
        
        Set obj_vendor = obj_md_dashboard.col_conditions_process_material(obj_kpi_on_time_pallet.str_material_vendor)
        Set obj_condition_process_material = obj_vendor.retrieve_condition_process_material( _
             obj_kpi_on_time_pallet.str_process_id, obj_kpi_on_time_pallet.str_process_version_id)
        Set obj_kpi_limit_source = obj_kpi_limit_mat_grp_wh.col_limits_source(obj_condition_process_material.str_condition_id)
        On Error GoTo 0
    Else
        Set obj_kpi_limit_source = obj_kpi_limit_mat_grp_wh.col_limits_source("default")
    End If
    
    On Error GoTo INFO_NOT_REGISTERED_CHECKPOINT
    Set retrieve_checkpoint = obj_kpi_limit_source.col_limits_checkpoint_place(obj_kpi_on_time_pallet.str_place_from)
    On Error GoTo 0
    
    Exit Function
WARN_NOT_FOUND_MATERIAL_GROUP_WH_RESOLVER:
    hndl_log.log db_log.TYPE_WARN, STR_CLASS_NAME, "retrieve_checkpoint", _
        "Material group wh resolver for BIN storage group:" & obj_kpi_on_time_pallet.str_material_bin_storage_group & _
        " and Vendor:" & obj_kpi_on_time_pallet.str_material_vendor & _
        " doesn't exist. Checkpoint:" & obj_kpi_on_time_pallet.str_place_to & " for process:" & obj_kpi_on_time_pallet.str_process_id & ">" & obj_kpi_on_time_pallet.str_process_version_id
    Set obj_kpi_limit_mat_grp_wh = obj_kpi_limit_process_version.col_limits_mat_grp_wh("default") ' # implement - change default string to constant
    bool_only_default = True
    Resume Next
    Exit Function
WARN_NOT_FOUND_CONDITION_PROCESS_MATERIAL:
    hndl_log.log db_log.TYPE_WARN, STR_CLASS_NAME, "retrieve_checkpoint", _
        "Condition for Process:" & obj_kpi_on_time_pallet.str_process_id & _
        ", Process version:" & obj_kpi_on_time_pallet.str_process_version_id & _
        " and Vendor:" & obj_kpi_on_time_pallet.str_material_vendor & _
        " doesn't exist. Checkpoint:" & obj_kpi_on_time_pallet.str_place_to & " for process:" & obj_kpi_on_time_pallet.str_process_id & ">" & obj_kpi_on_time_pallet.str_process_version_id
    Set obj_kpi_limit_source = obj_kpi_limit_mat_grp_wh.col_limits_source("default") ' # implement - change default string to constant
    bool_only_default = True
    Resume Next
    Exit Function
INFO_NOT_REGISTERED_CHECKPOINT:
    hndl_log.log db_log.TYPE_INFO, STR_CLASS_NAME, "retrieve_checkpoint", "Checkpoint:" & obj_kpi_on_time_pallet.str_place_to & " for process:" & obj_kpi_on_time_pallet.str_process_id & ">" & obj_kpi_on_time_pallet.str_process_version_id & ", for pallet: " & obj_kpi_on_time_pallet.str_id
End Function

Public Function retrieve_pallet_id_interval(obj_kpi_on_time_pallet As KPIOnTimePallet)
    retrieve_pallet_id_interval = obj_kpi_on_time_pallet.str_id & obj_kpi_on_time_pallet.str_place_from
End Function

' # method is called from MDLProcessDashboardListener
Public Function save(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_kpi_on_time_interval As KPIOnTimeInterval
    Dim obj_kpi_on_time_pallet As KPIOnTimePallet

    For Each obj_kpi_on_time_interval In col_kpi_intervals
        ' process all pallets in interval
        For Each obj_kpi_on_time_pallet In obj_kpi_on_time_interval.col_pallets_step
            If obj_kpi_on_time_pallet.str_id = "357020105501109173" Then
                Debug.Print "KPI interval" & obj_kpi_on_time_interval.str_id & ", Pallet data:" & obj_kpi_on_time_pallet.str_place_from & ">" & obj_kpi_on_time_pallet.str_place_to
                DoEvents
            End If
        
            obj_mdl_kpi_on_time_pallet.save_record_dynamic obj_kpi_on_time_pallet
        Next
    Next
    
    ' clean up
    'Set col_kpi_pallets = New Collection
'    Set col_kpi_intervals = New Collection
'    obj_mdl_kpi_on_time_pallet.close_data
End Function

Private Sub Class_Initialize()
    bool_db_data_process_unfinished_loaded = False
    
    Set col_kpi_intervals = New Collection
'    Set col_result_pallets = New Collection
    Set col_kpi_pallets = New Collection
    
    Set obj_checkpoint_type_resolver = New CheckpointTypeResolver ' # implement - cross check if this is the correct place where to initialize it
End Sub


