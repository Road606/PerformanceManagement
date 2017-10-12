Attribute VB_Name = "hndl_process"
Option Explicit

Public Const str_module As String = "hndl_process"
Public STR_DURATION_UNIT As String

Public obj_records As Collection
Public obj_records_closed As Collection

Public Function init()
    Set obj_records = New Collection
    
    STR_DURATION_UNIT = "n"
End Function

Public Function run(rg_record As Range)
    Dim obj_perf_record As Object

    If rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value = "357020105325325834" Or _
            rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value = "357020105330442410" _
        Then
        DoEvents
    End If

    Select Case rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TYPE_STARTED).Value
        Case db_transaction_type.STR_TRANSACTION_TYPE_PO_GR
            ' temporary solution skip pallets for processing
            If Not bin_stor_grp.is_processing(rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_STORAGE_GROUP_MATERIAL).Value) Then
                Set obj_perf_record = create_inbound(rg_record, db_process_subtype.STR_PO)
            End If
        Case db_transaction_type.STR_TRANSACTION_TYPE_PROD_ORD_GR
            If bin_stor_grp.is_outbound(rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_STORAGE_GROUP_MATERIAL).Value) Then
                Set obj_perf_record = create_outbound(rg_record, db_process_subtype.STR_PROD)
            Else
                Set obj_perf_record = create_inbound(rg_record, db_process_subtype.STR_PROD)
            End If
        Case db_transaction_type.STR_TRANSACTION_TYPE_TASK_LIST_CREATE
            Set obj_perf_record = create_supply(rg_record)
        Case db_transaction_type.STR_TRANSACTION_TYPE_BUILD_VHU, _
                db_transaction_type.STR_TRANSACTION_TYPE_HU_MOVE, _
                db_transaction_type.STR_TRANSACTION_TYPE_PROD_ORD_GI
            Set obj_perf_record = update(rg_record)
        Case db_transaction_type.STR_TRANSACTION_TYPE_PROD_ORD_DEPICK
            On Error GoTo INFO_DEPICKED_RECORD_NOT_FOUND
            obj_records.Remove rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_FROM).Value
            On Error GoTo 0
    End Select
    
    If Not obj_perf_record Is Nothing Then
        If obj_perf_record.is_closed Then
            On Error GoTo ERR_SAVE
            hndl_performance.save obj_perf_record
            obj_records.Remove obj_perf_record.str_pallet_id
            On Error GoTo 0
        ElseIf obj_perf_record.is_quality_checked Then
            obj_records.Remove obj_perf_record.str_pallet_id
        
            Set hndl_log.obj_log_record = New DBLogRecord
            hndl_log.obj_log_record.str_datetime = Now
            hndl_log.obj_log_record.str_type = db_log.TYPE_INFO
            hndl_log.obj_log_record.str_module = str_module
            hndl_log.obj_log_record.str_function = "run"
            hndl_log.obj_log_record.str_message = "Performance record out of scope> " & _
                rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value & ">" & _
                    rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_FROM).Value & ">" & _
                rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_CODE).Value
            hndl_log.save_record
        End If
    End If
    
    Exit Function
ERR_SAVE:
    Set hndl_log.obj_log_record = New DBLogRecord
    hndl_log.obj_log_record.str_datetime = Now
    hndl_log.obj_log_record.str_type = db_log.TYPE_INFO
    hndl_log.obj_log_record.str_module = str_module
    hndl_log.obj_log_record.str_function = "run"
    hndl_log.obj_log_record.str_message = "SAVE METHOD FAILED: " & _
        rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value & _
        rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value & _
        "Original message>" & Err.description
    hndl_log.save_record
INFO_DEPICKED_RECORD_NOT_FOUND:
    Set hndl_log.obj_log_record = New DBLogRecord
    hndl_log.obj_log_record.str_datetime = Now
    hndl_log.obj_log_record.str_type = db_log.TYPE_INFO
    hndl_log.obj_log_record.str_module = str_module
    hndl_log.obj_log_record.str_function = "run"
    hndl_log.obj_log_record.str_message = "Depicked pallet record not found> " & _
        rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value & ">" & _
            rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_FROM).Value & ">" & _
        rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_CODE).Value
    hndl_log.save_record
End Function

Public Function create_inbound(rg_record As Range, str_process_subtype As String) As Object
    If bin_stor_grp.is_hbw(rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_STORAGE_GROUP_MATERIAL).Value) Then
        Exit Function
    End If

    If str_process_subtype = db_process_subtype.STR_PO Then
        Set create_inbound = New PerformanceInboundPO
    ElseIf str_process_subtype = db_process_subtype.STR_PROD Then
        If bin_stor_grp.is_hbw(rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_STORAGE_GROUP_MATERIAL).Value) Then
            'Set create_inbound = New PerformanceInboundProdHBW
        Else
            Set create_inbound = New PerformanceInboundProd
        End If
    End If
        
    create_inbound.create_records rg_record
    On Error GoTo ERR_RECORD_DOUBLED
    obj_records.add create_inbound, create_inbound.str_pallet_id
    On Error GoTo 0
    Exit Function
ERR_RECORD_DOUBLED:
    If rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TYPE_FINISHED).Value <> _
            db_transaction_type.STR_TRANSACTION_TYPE_BUILD_VHU Then
        'Debug.Print str_module & ">run>ERR_RECORD_DOUBLED: " & rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value
        Set hndl_log.obj_log_record = New DBLogRecord
        hndl_log.obj_log_record.str_datetime = Now
        hndl_log.obj_log_record.str_type = db_log.TYPE_WARN
        hndl_log.obj_log_record.str_module = str_module
        hndl_log.obj_log_record.str_function = "create_inbound"
        hndl_log.obj_log_record.str_message = "RECORD_DOUBLED: " & _
            rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value & "->" & _
            rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
        hndl_log.save_record
    End If
    Exit Function
End Function

Public Function create_outbound(rg_record As Range, str_process_subtype As String) As Object

    If str_process_subtype = db_process_subtype.STR_PROD Then
        Set create_outbound = New PerformanceOutboundPush
    Else
        ' planned for deliveries
    End If
        
    create_outbound.create_records rg_record
    On Error GoTo ERR_RECORD_DOUBLED
    obj_records.add create_outbound, create_outbound.str_pallet_id
    On Error GoTo 0
    Exit Function
ERR_RECORD_DOUBLED:
    If rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_FINISHED).Value <> _
            db_transaction_type.STR_TRANSACTION_TYPE_BUILD_VHU Then
        'Debug.Print str_module & ">run>ERR_RECORD_DOUBLED: " & rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value
        Set hndl_log.obj_log_record = New DBLogRecord
        hndl_log.obj_log_record.str_datetime = Now
        hndl_log.obj_log_record.str_type = db_log.TYPE_WARN
        hndl_log.obj_log_record.str_module = str_module
        hndl_log.obj_log_record.str_function = "create_outbound"
        hndl_log.obj_log_record.str_message = "RECORD_DOUBLED: " & _
            rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value & _
            rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
        hndl_log.save_record
    End If
    Exit Function
End Function

Public Function create_supply(rg_record As Range) As Object

    If rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_CODE).Value = "" _
            And user.is_system(rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_USERNAME).Value) Then
        If rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TASK_LIST_TYPE).Value = db_task_list_type.STR_PICK Then
            ' autopick
            Set create_supply = New PerformanceSupplyAutopick
        Else
            ' replenishment
        End If
    Else
        ' other picking
    End If
            
    If Not create_supply Is Nothing Then
        create_supply.create_records rg_record
        On Error GoTo ERR_RECORD_DOUBLED
        obj_records.add create_supply, create_supply.str_pallet_id
        On Error GoTo 0
        End If
    Exit Function
ERR_RECORD_DOUBLED:
    If create_supply.str_process_subtype = db_process_subtype.STR_AUTO_PICK Then ' could happen if previous task list create record was deleted, but pallet wasn't issued to production
        obj_records.Remove create_supply.str_pallet_id
        obj_records.add create_supply, create_supply.str_pallet_id
    Else
        'Debug.Print str_module & ">run>ERR_RECORD_DOUBLED: " & rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value
        Set hndl_log.obj_log_record = New DBLogRecord
        hndl_log.obj_log_record.str_datetime = Now
        hndl_log.obj_log_record.str_type = db_log.TYPE_WARN
        hndl_log.obj_log_record.str_module = str_module
        hndl_log.obj_log_record.str_function = "create_supply"
        hndl_log.obj_log_record.str_message = "RECORD_DOUBLED: " & _
            rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value & _
            rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_STARTED).Value
        hndl_log.save_record
    End If
    Exit Function
End Function

Public Function update(rg_record As Range) As Object
    On Error GoTo ERR_RECORD_NOT_FOUND
    Set update = obj_records.Item(rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_FROM).Value)
    'obj_records.Remove rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value
    On Error GoTo 0
        
    ' change pallet_id if build vhu is incorporated
    If rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TYPE_FINISHED).Value = db_transaction_type.STR_TRANSACTION_TYPE_BUILD_VHU Or _
            rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_TRANSACTION_TYPE_STARTED).Value = db_transaction_type.STR_TRANSACTION_TYPE_BUILD_VHU _
            Then
        obj_records.Remove update.str_pallet_id
        update.str_pallet_id = rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_COMBI_VHU_TO).Value
        On Error GoTo INFO_PALLET_EXISTS
        obj_records.add update, update.str_pallet_id
        On Error GoTo 0
    End If
    
    update.update_records rg_record
    Exit Function
ERR_RECORD_NOT_FOUND:
    ' if record not found then it means it's a pallet for which GR or production request is missing
    ' for such a pallets we're not interested
    Exit Function
INFO_PALLET_EXISTS:
    Set update = obj_records.Item(update.str_pallet_id)
    Resume Next
    Exit Function
End Function
