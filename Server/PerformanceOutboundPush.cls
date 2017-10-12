VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceOutboundPush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Public obj_out_gate_push As PerformanceRecord
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
'Public bool_closed As Boolean
'Public bool_quality_checked As Boolean
'
'Private Sub Class_Initialize()
'    str_process_type = db_process_type.STR_OUTBOUND
'    str_process_subtype = db_process_subtype.STR_PROD
'
'    str_quality_check = db_quality.STR_CHECK_NO
'
'    bool_closed = False
'    bool_quality_checked = False
'
'    Set obj_out_gate_push = New PerformanceRecord
'    obj_out_gate_push.str_process_part = db_process_part.STR_GR_GATE_PUSH
'End Sub
'
'Public Function create_records(rg_record As Range)
'    Dim str_datetime As String
'    Dim str_bin As String
'    Dim str_building As String
'    Dim str_hall As String
'
'    ' init global
'    str_process_step = db_process_step_status.STR_GR_DONE
'    str_pallet_id = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value
'    str_material = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_MATERIAL).Value
'
'    increase_process_steps
'    find_out_quality_checked_status rg_record
'
'    ' init records
'    str_datetime = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'    str_building = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'    str_hall = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'    str_bin = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
'
'    ' receipt - gate push
'    obj_out_gate_push.open_obj
'    obj_out_gate_push.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'    obj_out_gate_push.str_start_bin = str_bin
'    obj_out_gate_push.str_start_building = str_building ' rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'    obj_out_gate_push.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'
'    ' check where received
'    If bin.is_production_hall(str_bin) Then
'        str_status = db_performance_inbound_status.STR_OK
'    ElseIf bin.is_gate_push(str_bin) Then
'        'str_status = db_performance_inbound_status.STR_GATE_PUSH ' bin to must be corrected in history app
'        str_status = db_performance_inbound_status.STR_OK
'
'        If rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_CODE) = db_transaction_code.STR_ZR14 Then
'            str_status = db_performance_inbound_status.STR_GATE_PUSH
'            str_status = str_status & ";" & db_performance_inbound_status.STR_ZR14
'            ' update start location data
'            'obj_out_gate_push.str_start_bin = str_bin ' rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_FROM).Value
'            'obj_out_gate_push.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_INBOUND_PLACE_BUILDING).Value
'            'obj_out_gate_push.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_INBOUND_PLACE_HALL).Value
'            ' update
'            update_records rg_record
'        End If
'    Else
'        str_status = db_performance_inbound_status.STR_GATE_PUSH
'    End If
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
'    find_out_quality_checked_status rg_record
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
'
'    If bin.is_gate_push(str_bin) Then
'        If str_process_step = db_process_step_status.STR_GR_DONE Then
'            If CDate(obj_out_gate_push.str_start_datetime) > CDate(str_datetime) Then
'                Exit Function
'            End If
'            ' receipt -> gate push
'            obj_out_gate_push.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
'            obj_out_gate_push.str_end_bin = str_bin
'            obj_out_gate_push.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'            obj_out_gate_push.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'            obj_out_gate_push.int_process_steps = int_process_steps
'            obj_out_gate_push.close_obj
'            ' update process step
'            str_process_step = db_process_step_status.STR_LOAD_PUSH_DONE
'
'            close_records
'        End If
'    End If
'End Function
'
'Public Function close_records()
'    If obj_out_gate_push.is_closed Then
'        obj_out_gate_push.str_material = str_material
'        obj_out_gate_push.str_pallet_id = str_pallet_id
'        obj_out_gate_push.str_process_type = str_process_type
'        obj_out_gate_push.str_process_subtype = str_process_subtype
'        obj_out_gate_push.str_status = str_status
'        obj_out_gate_push.str_quality_check = str_quality_check
'    End If
'
'    bool_closed = True
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
'Public Function is_closed() As Boolean
'    is_closed = bool_closed
'End Function
'
'Public Function is_quality_checked() As Boolean
'    is_quality_checked = bool_quality_checked
'End Function
'
'Public Function increase_process_steps()
'    int_process_steps = int_process_steps + 1
'End Function
'
'Public Function decrease_process_steps()
'    int_process_steps = int_process_steps - 1
'End Function
'
'Public Function get_start_datetime() As String
'    get_start_datetime = obj_out_gate_push.str_start_datetime
'End Function
