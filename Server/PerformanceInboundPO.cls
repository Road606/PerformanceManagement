VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceInboundPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Public obj_receipt_vna_rack As PerformanceRecord
'Public obj_receipt_vna_inbound As PerformanceRecord
'Public obj_vna_inbound_vna_rack As PerformanceRecord
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
'Public obj_prs As Collection
'
'Public bool_closed As Boolean
'Public bool_quality_checked As Boolean
'
'Private Sub Class_Initialize()
'    str_process_type = db_process_type.STR_INBOUND
'    str_process_subtype = db_process_subtype.STR_PO
'
'    str_quality_check = db_quality.STR_CHECK_NO
'
'    bool_closed = False
'    'bool_quality_checked = False
'
'
'    Set obj_prs = New Collection
'
'    Set obj_receipt_vna_rack = New PerformanceRecord
'    obj_receipt_vna_rack.str_process_part = db_process_part.STR_GR_VNA_RACK
'
'    Set obj_receipt_vna_inbound = New PerformanceRecord
'    obj_receipt_vna_inbound.str_process_part = db_process_part.STR_GR_VNA_INBOUND
'
'    Set obj_vna_inbound_vna_rack = New PerformanceRecord
'    obj_vna_inbound_vna_rack.str_process_part = db_process_part.STR_VNA_INBOUND_VNA_RACK
'
'    'Set obj_out_gate_push = New PerformanceRecord
'    'obj_out_gate_push.str_process_type = db_process_type.STR_PROD_HALL_GATE_PUSH
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
'    ' create records
'      ' receipt - vna rack
'    obj_receipt_vna_rack.open_obj
'    obj_receipt_vna_rack.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'    obj_receipt_vna_rack.str_start_bin = str_bin 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
'    obj_receipt_vna_rack.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'    obj_receipt_vna_rack.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'
'      ' receipt - vna inbound
'    obj_receipt_vna_inbound.open_obj
'    obj_receipt_vna_inbound.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
'    obj_receipt_vna_inbound.str_start_bin = str_bin
'    obj_receipt_vna_inbound.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'    obj_receipt_vna_inbound.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'
'    ' check where received
'    If bin.is_ra(str_bin) Then
'        str_status = db_performance_inbound_status.STR_OK
'    ElseIf bin.is_vna_bulk(str_bin) Then
'        str_status = db_performance_inbound_status.STR_VNA_INBOUND
'        update_records rg_record
'    ElseIf bin.is_vna_rack(str_bin) Then
'        str_status = db_performance_inbound_status.STR_VNA_RACK
'        'bool_closed = True  ' Set obj_pi = obj_close(rg_record, False) ' False says to use VHU_TO as pallet id
'        update_records rg_record
'    ElseIf bin.is_hbw_conveyor(str_bin) Then
'        str_status = db_performance_inbound_status.STR_HBW_CONVEYOR
'    Else
'        str_status = db_performance_inbound_status.STR_OTHERS
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
'    If bin.is_vna_bulk(rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value) Then
'        If Not obj_receipt_vna_inbound.is_closed Then
'            If CDate(obj_receipt_vna_inbound.str_start_datetime) > CDate(str_datetime) Then
'                Exit Function
'            End If
'
'            ' change process step
'            str_process_step = db_process_step_status.STR_VNA_INBOUND_DONE
'
'            ' receipt - vna inbound close
'            obj_receipt_vna_inbound.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
'            obj_receipt_vna_inbound.str_end_bin = str_bin 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
'            obj_receipt_vna_inbound.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'            obj_receipt_vna_inbound.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'            obj_receipt_vna_inbound.int_process_steps = int_process_steps - obj_vna_inbound_vna_rack.int_process_steps
'            obj_receipt_vna_inbound.close_obj
'
'            ' vna inbound -> vna rack open
'            obj_vna_inbound_vna_rack.open_obj
'            obj_vna_inbound_vna_rack.str_start_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
'            obj_vna_inbound_vna_rack.str_start_bin = str_bin 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
'            obj_vna_inbound_vna_rack.str_start_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'            obj_vna_inbound_vna_rack.str_start_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'            obj_vna_inbound_vna_rack.int_process_steps = int_process_steps
'        End If
'    ElseIf bin.is_vna_rack(rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value) Then
'        If Not obj_receipt_vna_rack.is_closed Then
'            ' vna inbound -> vna rack
'            If obj_vna_inbound_vna_rack.is_open Then
'                If CDate(obj_vna_inbound_vna_rack.str_start_datetime) > CDate(str_datetime) Then
'                    Exit Function
'                End If
'                obj_vna_inbound_vna_rack.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
'                obj_vna_inbound_vna_rack.str_end_bin = str_bin 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
'                obj_vna_inbound_vna_rack.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'                obj_vna_inbound_vna_rack.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'                obj_vna_inbound_vna_rack.int_process_steps = int_process_steps - obj_vna_inbound_vna_rack.int_process_steps
'                obj_vna_inbound_vna_rack.close_obj
'            End If
'
'            If CDate(obj_receipt_vna_rack.str_start_datetime) > CDate(str_datetime) Then
'                Exit Function
'            End If
'
'            ' change process step
'            str_process_step = db_process_step_status.STR_VNA_RACK_DONE
'
'            ' receipt - vna rack
'            obj_receipt_vna_rack.str_end_datetime = str_datetime 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value
'            obj_receipt_vna_rack.str_end_bin = str_bin 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO).Value
'            obj_receipt_vna_rack.str_end_building = str_building 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_BUILDING).Value
'            obj_receipt_vna_rack.str_end_hall = str_hall 'rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_OUTBOUND_PLACE_HALL).Value
'            obj_receipt_vna_rack.int_process_steps = int_process_steps
'            obj_receipt_vna_rack.close_obj
'
'            close_records
'        End If
'    End If
'End Function
'
'Public Function close_records()
'    If obj_receipt_vna_rack.is_closed Then
'        obj_receipt_vna_rack.str_material = str_material
'        obj_receipt_vna_rack.str_pallet_id = str_pallet_id
'        obj_receipt_vna_rack.str_process_type = str_process_type
'        obj_receipt_vna_rack.str_process_subtype = str_process_subtype
'        obj_receipt_vna_rack.str_status = str_status
'        obj_receipt_vna_rack.str_status = str_status
'        obj_receipt_vna_rack.str_quality_check = str_quality_check
'    End If
'
'    If obj_receipt_vna_inbound.is_closed Then
'        obj_receipt_vna_inbound.str_material = str_material
'        obj_receipt_vna_inbound.str_pallet_id = str_pallet_id
'        obj_receipt_vna_inbound.str_process_subtype = str_process_subtype
'        obj_receipt_vna_inbound.str_process_type = str_process_type
'        obj_receipt_vna_inbound.str_status = str_status
'        obj_receipt_vna_inbound.str_quality_check = str_quality_check
'    End If
'
'    If obj_vna_inbound_vna_rack.is_closed Then
'        obj_vna_inbound_vna_rack.str_material = str_material
'        obj_vna_inbound_vna_rack.str_pallet_id = str_pallet_id
'        obj_vna_inbound_vna_rack.str_process_subtype = str_process_subtype
'        obj_vna_inbound_vna_rack.str_process_type = str_process_type
'        obj_vna_inbound_vna_rack.str_status = str_status
'        obj_vna_inbound_vna_rack.str_quality_check = str_quality_check
'    End If
'
'    'If obj_out_gate_push.is_closed Then
'    '    obj_out_gate_push.str_material = str_material
'    '    obj_out_gate_push.str_pallet_id = str_pallet_id
'    '    obj_out_gate_push.str_process_subtype = str_process_subtype
'    'End If
'
'    bool_closed = True
'End Function
'
'Public Function is_closed() As Boolean
'    is_closed = bool_closed
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
'Public Function increase_process_steps()
'    int_process_steps = int_process_steps + 1
'End Function
'
'Public Function decrease_process_steps()
'    int_process_steps = int_process_steps - 1
'End Function
'
'Public Function get_start_datetime() As String
'    get_start_datetime = obj_receipt_vna_rack.str_start_datetime
'End Function
