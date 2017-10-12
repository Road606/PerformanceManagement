Attribute VB_Name = "hndl_performance"
Option Explicit

Public STR_WS_NAME As String
Public STR_DAILY_WS_NAME_KPI As String
Public STR_DAILY_WS_NAME_ADDITIONAL As String

Public STR_DATA_FIRST_CELL As String
Public STR_FIRST_ROW_TBL As String
Public STR_FIRST_ROW_DATA As String

Public LNG_CAPACITY_THRESHOLD As Long

Public rg_record As Range
Public ws As Worksheet
Public rg_last_row As Range

Public Function init()
    STR_WS_NAME = "data"
    STR_DATA_FIRST_CELL = "A2"
    
    STR_FIRST_ROW_TBL = "A1:U1"
    STR_FIRST_ROW_DATA = "A2:U2"
    
    LNG_CAPACITY_THRESHOLD = 2000
    
    Set ws = ThisWorkbook.Worksheets(STR_WS_NAME)
    Set rg_last_row = ws.Cells(ws.Range("A:A").CountLarge, 1)
End Function

Public Function clear()
    ws.UsedRange.Offset(1).clear
End Function

Public Function get_data_daily_kpi(Optional str_date As String = "", Optional bool_closed_only = True) As Range
    Dim ws_local As Worksheet
    Dim rg_tbl As Range
    'Dim rg_data As Range
    
    Set ws_local = ThisWorkbook.Worksheets(STR_DAILY_WS_NAME_KPI)
    ws_local.Activate
    Set rg_tbl = ws_local.Range(STR_FIRST_ROW_TBL)
    Set rg_tbl = Range(rg_tbl, rg_tbl.End(xlDown).End(xlToRight))
    'rg_tbl.AutoFilter _
    '    Field:=db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STATUS + 1, _
    '    Criteria1:=STR_TRANSACTION_STATUS_CLOSED
'    If str_date <> "" Then
'        rg_tbl.AutoFilter _
'        Field:=db_performance.INT_OFFSET_TRANSACTION_DATE + 1, _
'        Criteria1:=str_date
'    End If
        
    Set get_data_daily_kpi = ws_local.Range(STR_FIRST_ROW_DATA)
    If get_data_daily_kpi.End(xlDown).Value = "" Then
        Set get_data_daily_kpi = Nothing
    Else
        Set get_data_daily_kpi = Range(get_data_daily_kpi, get_data_daily_kpi.End(xlDown)).SpecialCells(xlCellTypeVisible)
    End If
    
    
    rg_tbl.AutoFilter ' clear filter
    rg_tbl.AutoFilter ' add filter, because it was removed in previous command
End Function

Public Function get_data(Optional str_date As String = "", Optional bool_closed_only = True) As Range
    Dim ws_local As Worksheet
    Dim rg_tbl As Range
    'Dim rg_data As Range
    
    Set ws_local = ThisWorkbook.Worksheets(STR_WS_NAME)
    ws_local.Activate
    Set rg_tbl = ws_local.Range(STR_FIRST_ROW_TBL)
    Set rg_tbl = Range(rg_tbl, rg_tbl.End(xlDown).End(xlToRight))
    'rg_tbl.AutoFilter _
    '    Field:=db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STATUS + 1, _
    '    Criteria1:=STR_TRANSACTION_STATUS_CLOSED
    If str_date <> "" Then
        rg_tbl.AutoFilter _
        Field:=db_performance.INT_OFFSET_TRANSACTION_DATE + 1, _
        Criteria1:=str_date
    End If
        
    Set get_data = ws_local.Range(STR_FIRST_ROW_DATA)
    If get_data.End(xlDown).Value = "" Then
        Set get_data = Nothing
    Else
        Set get_data = Range(get_data, get_data.End(xlDown)).SpecialCells(xlCellTypeVisible)
    End If
    
    
    rg_tbl.AutoFilter ' clear filter
    rg_tbl.AutoFilter ' add filter, because it was removed in previous command
End Function

Public Function save(obj_perf_record As Object)
    Select Case obj_perf_record.str_process_type
        Case db_process_type.STR_INBOUND
            If obj_perf_record.obj_receipt_vna_rack.is_closed Then
                save_obj obj_perf_record.obj_receipt_vna_rack
            End If
            
            If obj_perf_record.obj_receipt_vna_inbound.is_closed Then
                save_obj obj_perf_record.obj_receipt_vna_inbound
            End If
            
            If obj_perf_record.obj_vna_inbound_vna_rack.is_closed Then
                save_obj obj_perf_record.obj_vna_inbound_vna_rack
            End If
        Case db_process_type.STR_OUTBOUND
            If obj_perf_record.obj_out_gate_push.is_closed Then
                save_obj obj_perf_record.obj_out_gate_push
            End If
        Case db_process_type.STR_SUPPLY
            If obj_perf_record.obj_vna_rack_production.is_closed Then
                save_obj obj_perf_record.obj_vna_rack_production
            End If
            
            If obj_perf_record.obj_vna_rack_vna_inbound.is_closed Then
                save_obj obj_perf_record.obj_vna_rack_vna_inbound
            End If
                
            If obj_perf_record.obj_vna_inbound_ta_inbound.is_closed Then
                save_obj obj_perf_record.obj_vna_inbound_ta_inbound
            End If
            
            If obj_perf_record.obj_vna_inbound_ta_rack.is_closed Then
                save_obj obj_perf_record.obj_vna_inbound_ta_rack
            End If
            
            If obj_perf_record.obj_vna_inbound_production.is_closed Then
                save_obj obj_perf_record.obj_vna_inbound_production
            End If
            
            If obj_perf_record.obj_ta_inbound_ta_rack.is_closed Then
                save_obj obj_perf_record.obj_ta_inbound_ta_rack
            End If
            
            If obj_perf_record.obj_ta_inbound_production.is_closed Then
                save_obj obj_perf_record.obj_ta_inbound_production
            End If
            
            If obj_perf_record.obj_ta_rack_production.is_closed Then
                save_obj obj_perf_record.obj_ta_rack_production
            End If
    End Select
End Function

Public Function save_obj(obj_pir As PerformanceRecord)
    Dim rg_record As Range
    
    Set rg_record = rg_last_row.End(xlUp).Offset(1)
    
    rg_record.Offset(0, db_performance.INT_OFFSET_PALLET).NumberFormat = "@"
    rg_record.Offset(0, db_performance.INT_OFFSET_PALLET).Value = obj_pir.str_pallet_id
    rg_record.Offset(0, db_performance.INT_OFFSET_MATERIAL).Value = obj_pir.str_material
    rg_record.Offset(0, db_performance.INT_OFFSET_START_DATETIME).Value = obj_pir.str_start_datetime
    rg_record.Offset(0, db_performance.INT_OFFSET_END_DATETIME).Value = obj_pir.str_end_datetime
    rg_record.Offset(0, db_performance.INT_OFFSET_DURATION).Value = obj_pir.lng_duration
    rg_record.Offset(0, db_performance.INT_OFFSET_START_BIN).Value = obj_pir.str_start_bin
    rg_record.Offset(0, db_performance.INT_OFFSET_END_BIN).Value = obj_pir.str_end_bin
    rg_record.Offset(0, db_performance.INT_OFFSET_START_BUILDING).Value = obj_pir.str_start_building
    rg_record.Offset(0, db_performance.INT_OFFSET_END_BUILDING).Value = obj_pir.str_end_building
    rg_record.Offset(0, db_performance.INT_OFFSET_START_HALL).Value = obj_pir.str_start_hall
    rg_record.Offset(0, db_performance.INT_OFFSET_END_HALL).Value = obj_pir.str_end_hall
    On Error GoTo ERR_NO_DATE
    rg_record.Offset(0, db_performance.INT_OFFSET_TRANSACTION_DATE).Value = Format(DateValue(obj_pir.str_start_datetime), "dd.mm.yyyy")
    rg_record.Offset(0, db_performance.INT_OFFSET_TRANSACTION_HOUR).Value = Hour(obj_pir.str_start_datetime)
    rg_record.Offset(0, db_performance.INT_OFFSET_TRANSACTION_WEEKDAY).Value = Weekday(obj_pir.str_start_datetime, vbMonday)
    rg_record.Offset(0, db_performance.INT_OFFSET_TRANSACTION_SHIFT).Value = hndl_master_shift.find_shift( _
        rg_record.Offset(0, db_performance.INT_OFFSET_TRANSACTION_DATE).Value & "-" & rg_record.Offset(0, db_performance.INT_OFFSET_TRANSACTION_HOUR))
    On Error GoTo 0
    rg_record.Offset(0, db_performance.INT_OFFSET_PROCESS_TYPE).Value = obj_pir.str_process_type
    rg_record.Offset(0, db_performance.INT_OFFSET_PROCESS_SUBTYPE).Value = obj_pir.str_process_subtype
    rg_record.Offset(0, db_performance.INT_OFFSET_PROCESS_PART).Value = obj_pir.str_process_part
    rg_record.Offset(0, db_performance.INT_OFFSET_PROCESS_STEP).Value = obj_pir.int_process_steps
    rg_record.Offset(0, db_performance.INT_OFFSET_PROCESS_STATUS).Value = obj_pir.str_status
    
    rg_record.Offset(0, db_performance.INT_OFFSET_QUALITY_CHECK).Value = obj_pir.str_quality_check
    Exit Function
ERR_NO_DATE:
    Resume Next
End Function

Public Function is_capacity_above_threshold() As Boolean
    'Dim ws As Worksheet
    Dim rg As Range
    
    'Set ws = ThisWorkbook.Worksheets(STR_WS_NAME)
    'ws.Activate
    Set rg = ws.Cells(ws.Range("A:A").CountLarge, 1).End(xlUp)
    
    If rg.Row > LNG_CAPACITY_THRESHOLD Then
        is_capacity_above_threshold = True
    End If
End Function
