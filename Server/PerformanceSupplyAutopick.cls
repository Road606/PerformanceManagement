VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceSupplyAutopick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Public obj_vna_rack_production As PerformanceRecord
'Public obj_vna_rack_vna_inbound As PerformanceRecord
'Public obj_vna_inbound_ta_inbound As PerformanceRecord
'Public obj_vna_inbound_ta_rack As PerformanceRecord
'Public obj_vna_inbound_production As PerformanceRecord
'Public obj_ta_inbound_ta_rack As PerformanceRecord
'Public obj_ta_inbound_production As PerformanceRecord
'Public obj_ta_rack_production As PerformanceRecord
'
'Public str_material As String
'Public str_pallet_id As String
'
'Public str_process_type As String
'Public str_process_subtype As String
'Public str_status As String
'
'Public str_process_step As String
'Public int_process_steps As Integer
'
'Public str_quality_check As String
'
'Public BOOL_CLOSED As Boolean
'Public bool_quality_checked As Boolean
'
'Private Sub Class_Initialize()
'    str_process_type = db_process_type.STR_SUPPLY
'    str_process_subtype = db_process_subtype.STR_AUTO_PICK
'
'    Set obj_vna_rack_production = New PerformanceRecord
'    obj_vna_rack_production.str_process_part = db_process_part.STR_VNA_RACK_ISSUE
'
'    Set obj_vna_rack_vna_inbound = New PerformanceRecord
'    obj_vna_rack_vna_inbound.str_process_part = db_process_part.STR_VNA_RACK_VNA_INBOUND
'
'    Set obj_vna_inbound_ta_inbound = New PerformanceRecord
'    obj_vna_inbound_ta_inbound.str_process_part = db_process_part.STR_VNA_INBOUND_TA_INBOUND
'
'    Set obj_vna_inbound_ta_rack = New PerformanceRecord
'    obj_vna_inbound_ta_rack.str_process_part = db_process_part.STR_VNA_INBOUND_TA_RACK
'
'    Set obj_vna_inbound_production = New PerformanceRecord
'    obj_vna_inbound_production.str_process_part = db_process_part.STR_VNA_INBOUND_ISSUE
'
'    Set obj_ta_inbound_ta_rack = New PerformanceRecord
'    obj_ta_inbound_ta_rack.str_process_part = db_process_part.STR_TA_INBOUND_TA_RACK
'
'    Set obj_ta_inbound_production = New PerformanceRecord
'    obj_ta_inbound_production.str_process_part = db_process_part.STR_TA_INBOUND_ISSUE
'
'    Set obj_ta_rack_production = New PerformanceRecord
'    obj_ta_rack_production.str_process_part = db_process_part.STR_TA_RACK_ISSUE
'
'    BOOL_CLOSED = False
'    bool_quality_checked = False
'
'    str_quality_check = db_quality.STR_CHECK_NO
'End Sub
'
'Public Function create_records(rg_record As Range)
'    Dim str_datetime As String
'    Dim str_bin As String
'    Dim str_building As String
'    Dim str_hall As String
'
'    ' init global
'    str_process_step = db_process_step_status.STR_VNA_RACK_DONE
'    str_pallet_id = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value
'    str_material = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_MATERIAL).Value
'
'    increase_process_steps
'    'find_out_quality_checked_status rg_record
'
'    ' init records
'    str_datetime = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'    str_building = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_INBOUND_PLACE_BUILDING).Value
'    str_hall = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_INBOUND_PLACE_HALL).Value
'    str_bin = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_FROM).Value
'
'
'    ' create records
'      ' vna rack - production
'    obj_vna_rack_production.open_obj
'    obj_vna_rack_production.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'    obj_vna_rack_production.str_start_bin = str_bin 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
'    obj_vna_rack_production.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'    obj_vna_rack_production.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'    obj_vna_rack_production.int_process_steps = int_process_steps
'
'      ' vna rack - vna inbound
'    obj_vna_rack_vna_inbound.open_obj
'    obj_vna_rack_vna_inbound.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'    obj_vna_rack_vna_inbound.str_start_bin = str_bin 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
'    obj_vna_rack_vna_inbound.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'    obj_vna_rack_vna_inbound.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'    obj_vna_rack_vna_inbound.int_process_steps = int_process_steps
'End Function
'
'Public Function update_records(rg_record As Range)
'    Dim str_datetime As String
'    Dim str_bin As String
'    Dim str_building As String
'    Dim str_hall As String
'
'    ' global
'    increase_process_steps
'    'find_out_quality_checked_status rg_record
'
'    ' init records
'    If rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TYPE_FINISHED).Value = "" Then
'        str_datetime = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'    Else
'        str_datetime = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
'    End If
'
'    str_bin = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
'    str_building = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'    str_hall = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'      ' since 11/2015 task list create doesn't contain material number, error will be corrected in COMBI
'    str_material = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_MATERIAL).Value
'
'    ' vna rack -> vna inbound
'    If bin.is_vna_bulk(str_bin) Then
'        If Not str_process_step = db_process_step_status.STR_VNA_INBOUND_DONE Then
'            ' vna rack -> vna inbound
'            If obj_vna_rack_vna_inbound.is_open Then
'                If CDate(obj_vna_rack_vna_inbound.str_start_datetime) > CDate(str_datetime) Then
'                    Exit Function
'                End If
'                obj_vna_rack_vna_inbound.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
'                obj_vna_rack_vna_inbound.str_end_bin = str_bin
'                obj_vna_rack_vna_inbound.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'                obj_vna_rack_vna_inbound.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'                obj_vna_rack_vna_inbound.int_process_steps = int_process_steps - obj_vna_rack_vna_inbound.int_process_steps
'                obj_vna_rack_vna_inbound.close_obj
'            End If
'
'            ' open new
'              ' vna inbound -> ta inbound
'            obj_vna_inbound_ta_inbound.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'            obj_vna_inbound_ta_inbound.str_start_bin = str_bin
'            obj_vna_inbound_ta_inbound.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'            obj_vna_inbound_ta_inbound.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'            obj_vna_inbound_ta_inbound.int_process_steps = int_process_steps
'            obj_vna_inbound_ta_inbound.open_obj
'
'              ' vna inbound -> ta rack
'            obj_vna_inbound_ta_rack.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'            obj_vna_inbound_ta_rack.str_start_bin = str_bin
'            obj_vna_inbound_ta_rack.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'            obj_vna_inbound_ta_rack.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'            obj_vna_inbound_ta_rack.int_process_steps = int_process_steps
'            obj_vna_inbound_ta_rack.open_obj
'
'              ' vna inbound -> production
'            obj_vna_inbound_production.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'            obj_vna_inbound_production.str_start_bin = str_bin
'            obj_vna_inbound_production.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'            obj_vna_inbound_production.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'            obj_vna_inbound_production.int_process_steps = int_process_steps
'            obj_vna_inbound_production.open_obj
'
'            ' update process step
'            str_process_step = db_process_step_status.STR_VNA_INBOUND_DONE
'        End If
'    ElseIf bin.is_ta_bulk(str_bin) Then
'        If Not str_process_step = db_process_step_status.STR_TA_INBOUND_DONE Then
'            ' vna inbound -> ta inbound
'            If obj_vna_inbound_ta_inbound.is_open Then
'                If str_pallet_id = "357020105330442410" Then
'                    DoEvents
'                End If
'
'                If CDate(obj_vna_inbound_ta_inbound.str_start_datetime) > CDate(str_datetime) Then
'                    Exit Function
'                End If
'
'                obj_vna_inbound_ta_inbound.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
'                obj_vna_inbound_ta_inbound.str_end_bin = str_bin
'                obj_vna_inbound_ta_inbound.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'                obj_vna_inbound_ta_inbound.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'                obj_vna_inbound_ta_inbound.int_process_steps = int_process_steps - obj_vna_inbound_ta_inbound.int_process_steps
'                obj_vna_inbound_ta_inbound.close_obj
'            End If
'
'            ' open new
'              ' ta inbound -> ta rack
'            obj_ta_inbound_ta_rack.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'            obj_ta_inbound_ta_rack.str_start_bin = str_bin
'            obj_ta_inbound_ta_rack.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'            obj_ta_inbound_ta_rack.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'            obj_ta_inbound_ta_rack.int_process_steps = int_process_steps
'            obj_ta_inbound_ta_rack.open_obj
'
'              ' ta inbound -> production
'            obj_ta_inbound_production.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'            obj_ta_inbound_production.str_start_bin = str_bin
'            obj_ta_inbound_production.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'            obj_ta_inbound_production.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'            obj_ta_inbound_production.int_process_steps = int_process_steps
'            obj_ta_inbound_production.open_obj
'
'            ' update process step
'            str_process_step = db_process_step_status.STR_TA_INBOUND_DONE
'        End If
'    ElseIf bin.is_ta_rack(str_bin) Then
'        If Not str_process_step = db_process_step_status.STR_TA_RACK_DONE Then
'            ' vna inbound -> ta rack
'            If obj_vna_inbound_ta_rack.is_open Then
'                If str_pallet_id = "357020105325325834" Then
'                    DoEvents
'                End If
'
'                If CDate(obj_vna_inbound_ta_rack.str_start_datetime) > CDate(str_datetime) Then
'                    Exit Function
'                End If
'
'                obj_vna_inbound_ta_rack.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
'                obj_vna_inbound_ta_rack.str_end_bin = str_bin
'                obj_vna_inbound_ta_rack.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'                obj_vna_inbound_ta_rack.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'                obj_vna_inbound_ta_rack.int_process_steps = int_process_steps - obj_vna_inbound_ta_rack.int_process_steps
'                obj_vna_inbound_ta_rack.close_obj
'            End If
'
'            ' ta inbound -> ta rack
'            If obj_ta_inbound_ta_rack.is_open Then
'                If CDate(obj_ta_inbound_ta_rack.str_start_datetime) > CDate(str_datetime) Then
'                    Exit Function
'                End If
'
'                obj_ta_inbound_ta_rack.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
'                obj_ta_inbound_ta_rack.str_end_bin = str_bin
'                obj_ta_inbound_ta_rack.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'                obj_ta_inbound_ta_rack.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'                obj_ta_inbound_ta_rack.int_process_steps = int_process_steps - obj_ta_inbound_ta_rack.int_process_steps
'                obj_ta_inbound_ta_rack.close_obj
'            End If
'
'            ' open new
'              ' ta rack -> production
'            obj_ta_rack_production.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'            obj_ta_rack_production.str_start_bin = str_bin
'            obj_ta_rack_production.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'            obj_ta_rack_production.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'            obj_ta_rack_production.int_process_steps = int_process_steps
'            obj_ta_rack_production.open_obj
'
'            ' update process step
'            str_process_step = db_process_step_status.STR_TA_RACK_DONE
'        End If
'    ElseIf bin.is_production_line(str_bin) Then
'        If Not str_process_step = db_process_step_status.STR_SUPPLY_DONE Then
'            ' close open movements
'              ' vna inbound -> production - bin.is_vna_bulk(str_bin)
'            If obj_vna_inbound_production.is_open Then
'                If CDate(obj_vna_inbound_production.str_start_datetime) > CDate(str_datetime) Then
'                    Exit Function
'                End If
'
'                obj_vna_inbound_production.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'                obj_vna_inbound_production.str_end_bin = str_bin
'                obj_vna_inbound_production.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'                obj_vna_inbound_production.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'                obj_vna_inbound_production.int_process_steps = int_process_steps - obj_vna_inbound_production.int_process_steps
'                obj_vna_inbound_production.close_obj
'            End If
'
'              ' ta inbound -> production - open in bin.is_ta_bulk(str_bin)
'            If obj_ta_inbound_production.is_open Then
'                If CDate(obj_ta_inbound_production.str_start_datetime) > CDate(str_datetime) Then
'                    Exit Function
'                End If
'
'                obj_ta_inbound_production.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'                obj_ta_inbound_production.str_end_bin = str_bin
'                obj_ta_inbound_production.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'                obj_ta_inbound_production.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'                obj_ta_inbound_production.int_process_steps = int_process_steps - obj_ta_inbound_production.int_process_steps
'                obj_ta_inbound_production.close_obj
'            End If
'
'              ' ta rack -> production - open in bin.is_ta_rack(str_bin)
'            If obj_ta_rack_production.is_open Then
'                If CDate(obj_ta_rack_production.str_start_datetime) > CDate(str_datetime) Then
'                    Exit Function
'                End If
'
'                obj_ta_rack_production.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'                obj_ta_rack_production.str_end_bin = str_bin
'                obj_ta_rack_production.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'                obj_ta_rack_production.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'                obj_ta_rack_production.int_process_steps = int_process_steps - obj_ta_rack_production.int_process_steps
'                obj_ta_rack_production.close_obj
'            End If
'
'              ' vna rack -> production - open in create records
'            obj_vna_rack_production.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'            obj_vna_rack_production.str_end_bin = str_bin 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
'            obj_vna_rack_production.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'            obj_vna_rack_production.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'            obj_vna_rack_production.int_process_steps = int_process_steps - obj_vna_rack_production.int_process_steps
'            obj_vna_rack_production.close_obj
'
'            ' update process step
'            str_process_step = db_process_step_status.STR_SUPPLY_DONE
'
'            close_records
'        End If
'    End If
'End Function
'
'
'Public Function increase_process_steps()
'    int_process_steps = int_process_steps + 1
'End Function
'
'Public Function decrease_process_steps()
'    int_process_steps = int_process_steps - 1
'End Function
'
'Public Function close_records()
'    ' vna rack - production
'    obj_vna_rack_production.str_material = str_material
'    obj_vna_rack_production.str_pallet_id = str_pallet_id
'    obj_vna_rack_production.str_process_type = str_process_type
'    obj_vna_rack_production.str_process_subtype = str_process_subtype
'    obj_vna_rack_production.str_status = str_status
'
'    If obj_vna_rack_production.is_closed Then
'        obj_vna_rack_production.str_status = db_performance_supply_status.STR_OK
'    Else
'        obj_vna_rack_production.str_status = db_performance_supply_status.STR_SKIPPED
'    End If
'
'    ' vna rack - vna inbound
'    obj_vna_rack_vna_inbound.str_material = str_material
'    obj_vna_rack_vna_inbound.str_pallet_id = str_pallet_id
'    obj_vna_rack_vna_inbound.str_process_type = str_process_type
'    obj_vna_rack_vna_inbound.str_process_subtype = str_process_subtype
'    obj_vna_rack_vna_inbound.str_status = str_status
'
'    If obj_vna_rack_vna_inbound.is_closed Then
'        obj_vna_rack_vna_inbound.str_status = db_performance_supply_status.STR_OK
'    Else
'        obj_vna_rack_vna_inbound.str_status = db_performance_supply_status.STR_SKIPPED
'        obj_vna_rack_vna_inbound.str_start_datetime = obj_vna_rack_production.str_start_datetime
'    End If
'
'    ' vna inbound - ta inbound
'    obj_vna_inbound_ta_inbound.str_material = str_material
'    obj_vna_inbound_ta_inbound.str_pallet_id = str_pallet_id
'    obj_vna_inbound_ta_inbound.str_process_type = str_process_type
'    obj_vna_inbound_ta_inbound.str_process_subtype = str_process_subtype
'    obj_vna_inbound_ta_inbound.str_status = str_status
'
'    If obj_vna_inbound_ta_inbound.is_closed Then
'        obj_vna_inbound_ta_inbound.str_status = db_performance_supply_status.STR_OK
'    Else
'        obj_vna_inbound_ta_inbound.str_status = db_performance_supply_status.STR_SKIPPED
'        obj_vna_inbound_ta_inbound.str_start_datetime = obj_vna_rack_production.str_start_datetime
'    End If
'
'    ' vna inbound - ta rack
'    obj_vna_inbound_ta_rack.str_material = str_material
'    obj_vna_inbound_ta_rack.str_pallet_id = str_pallet_id
'    obj_vna_inbound_ta_rack.str_process_type = str_process_type
'    obj_vna_inbound_ta_rack.str_process_subtype = str_process_subtype
'    obj_vna_inbound_ta_rack.str_status = str_status
'
'    If obj_vna_inbound_ta_rack.is_closed Then
'        obj_vna_inbound_ta_rack.str_status = db_performance_supply_status.STR_OK
'    Else
'        obj_vna_inbound_ta_rack.str_status = db_performance_supply_status.STR_SKIPPED
'        obj_vna_inbound_ta_rack.str_start_datetime = obj_vna_rack_production.str_start_datetime
'    End If
'
'    ' vna inbound - production
'    obj_vna_inbound_production.str_material = str_material
'    obj_vna_inbound_production.str_pallet_id = str_pallet_id
'    obj_vna_inbound_production.str_process_type = str_process_type
'    obj_vna_inbound_production.str_process_subtype = str_process_subtype
'    obj_vna_inbound_production.str_status = str_status
'
'    If obj_vna_inbound_production.is_closed Then
'        obj_vna_inbound_production.str_status = db_performance_supply_status.STR_OK
'    Else
'        obj_vna_inbound_production.str_status = db_performance_supply_status.STR_SKIPPED
'        obj_vna_inbound_production.str_start_datetime = obj_vna_rack_production.str_start_datetime
'    End If
'
'    ' ta inbound - ta rack
'    obj_ta_inbound_ta_rack.str_material = str_material
'    obj_ta_inbound_ta_rack.str_pallet_id = str_pallet_id
'    obj_ta_inbound_ta_rack.str_process_type = str_process_type
'    obj_ta_inbound_ta_rack.str_process_subtype = str_process_subtype
'    obj_ta_inbound_ta_rack.str_status = str_status
'
'    If obj_ta_inbound_ta_rack.is_closed Then
'        obj_ta_inbound_ta_rack.str_status = db_performance_supply_status.STR_OK
'    Else
'        obj_ta_inbound_ta_rack.str_status = db_performance_supply_status.STR_SKIPPED
'        obj_ta_inbound_ta_rack.str_start_datetime = obj_vna_rack_production.str_start_datetime
'    End If
'
'    ' ta inbound - production
'    obj_ta_inbound_production.str_material = str_material
'    obj_ta_inbound_production.str_pallet_id = str_pallet_id
'    obj_ta_inbound_production.str_process_type = str_process_type
'    obj_ta_inbound_production.str_process_subtype = str_process_subtype
'    obj_ta_inbound_production.str_status = str_status
'
'    If obj_ta_inbound_production.is_closed Then
'        obj_ta_inbound_production.str_status = db_performance_supply_status.STR_OK
'    Else
'        obj_ta_inbound_production.str_status = db_performance_supply_status.STR_SKIPPED
'        obj_ta_inbound_production.str_start_datetime = obj_vna_rack_production.str_start_datetime
'    End If
'
'    ' ta rack - production
'    obj_ta_rack_production.str_material = str_material
'    obj_ta_rack_production.str_pallet_id = str_pallet_id
'    obj_ta_rack_production.str_process_type = str_process_type
'    obj_ta_rack_production.str_process_subtype = str_process_subtype
'    obj_ta_rack_production.str_status = str_status
'
'    If obj_ta_rack_production.is_closed Then
'        obj_ta_rack_production.str_status = db_performance_supply_status.STR_OK
'    Else
'        obj_ta_rack_production.str_status = db_performance_supply_status.STR_SKIPPED
'        obj_ta_rack_production.str_start_datetime = obj_vna_rack_production.str_start_datetime
'    End If
'
'    BOOL_CLOSED = True
'End Function
'
'Public Function is_closed() As Boolean
'    is_closed = BOOL_CLOSED
'End Function
'
'Public Function find_out_quality_checked_status(rg_record As Range)
'    Dim str_stock_type As String
'
'    If str_quality_check = db_quality.STR_CHECK_NO Then
'
'        str_stock_type = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_STOCK_TYPE_FROM).Value
'
'        If str_stock_type = "" Then
'            str_stock_type = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_STOCK_TYPE_TO).Value
'        End If
'
'        If str_stock_type = db_quality.STR_STOCK_TYPE_INSPECTION Or str_stock_type = db_quality.STR_STOCK_TYPE_BLOCKED Then
'            'bool_quality_checked = True
'            str_quality_check = db_quality.STR_CHECK_YES
'        End If
'    End If
'End Function
'
'Public Function is_quality_checked() As Boolean
'    is_quality_checked = bool_quality_checked
'End Function
'
'Public Function get_start_datetime() As String
'    get_start_datetime = obj_vna_rack_production.str_start_datetime
'End Function
