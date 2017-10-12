Attribute VB_Name = "new_ctrl_pallet"
Option Explicit

Public Const str_module As String = "new_ctrl_pallet"

' input
Public obj_data_provider_info As FileExcelDataProviderInfo

' output
Public obj_mdl_finished As MDLDataProcessExcel
Public obj_mdl_unfinished As MDLDataProcessExcel

Public col_pallets As Collection

Public Function init()
    Set col_pallets = New Collection
End Function

Public Function process_record(obj_record As DBHistoryRecord)
    Dim obj_pallet As Pallet
    Dim byte_record_action As Byte

    If obj_record.str_combi_vhu_from = "357020105503013287" And obj_record.str_user_name = "CZVITSKO" Then
        DoEvents
    End If
'    Select Case obj_record.str_combi_vhu_from
'        Case "357020105467288981"
'            DoEvents
'    End Select

    On Error GoTo INFO_NEW_PALLET
    Set obj_pallet = get_pallet(obj_record.str_combi_vhu_from) 'get_pallet_from_record(obj_record)
    On Error GoTo 0
        
            
    ' based on process master version check what to do with current transaction
    ' #implement UPDATE/CLOSE/DELETE
    ' evaluate all statuses for current record
    If new_ctrl_process.is_transaciton_valid_for_action(new_ctrl_process_master_action.STR_UPDATE, obj_record, obj_pallet.obj_process) Then
        byte_record_action = byte_record_action + new_ctrl_process_master_action.BYTE_UPDATE
    End If
    
    If obj_record.str_combi_vhu_from = "357020105502987503" Then
        DoEvents
    End If
    If new_ctrl_process.is_transaciton_valid_for_action(new_ctrl_process_master_action.STR_CLOSE, obj_record, obj_pallet.obj_process) Then
        byte_record_action = byte_record_action + new_ctrl_process_master_action.BYTE_CLOSE
    End If
    
    If new_ctrl_process.is_transaciton_valid_for_action(new_ctrl_process_master_action.STR_DELETE, obj_record, obj_pallet.obj_process) Then
        byte_record_action = byte_record_action + new_ctrl_process_master_action.BYTE_DELETE
    End If
    
    If obj_pallet.str_id = "357020105503005961" Then
        DoEvents
    End If
    
    'If obj_pallet.obj_process.obj_master_version.obj_master.get_action(new_ctrl_process_master_action.STR_UDATE).
    If byte_record_action And new_ctrl_process_master_action.BYTE_UPDATE Then
        If Not swap_pallet_id(obj_pallet, obj_record) Then ' if pallet is not aggregated during pallet id swap then update pallet
            update_record obj_record, obj_pallet, byte_record_action
        End If
    End If
    
    If byte_record_action And new_ctrl_process_master_action.BYTE_CLOSE Then
        close_record obj_record, obj_pallet
    End If
    
    If byte_record_action And new_ctrl_process_master_action.BYTE_DELETE Then
        delete_record obj_record, obj_pallet
    End If
    
'    If new_ctrl_process.is_transaciton_valid_for_action(new_ctrl_process_master_action.STR_UDATE, obj_record, obj_pallet.obj_process) Then
'        update_record obj_record, obj_pallet
'    End If
'
'    If new_ctrl_process.is_transaciton_valid_for_action(new_ctrl_process_master_action.STR_CLOSE, obj_record, obj_pallet.obj_process) Then
'        close_record obj_record, obj_pallet
'    End If
'
'    If new_ctrl_process.is_transaciton_valid_for_action(new_ctrl_process_master_action.STR_DELETE, obj_record, obj_pallet.obj_process) Then
'        delete_record obj_record, obj_pallet
'    End If
    
'    If obj_pallet.obj_process.byte_status = new_ctrl_process.BYTE_STATUS_CLOSED Then
'        Debug.Print "Writing process to db"
'        save_closed_pallet obj_pallet
'    End If
    Exit Function
INFO_NEW_PALLET:
    Set obj_pallet = create_record(obj_record)
    If Not obj_pallet Is Nothing Then
        Resume Next
    End If
End Function

Public Function create_record(obj_record As DBHistoryRecord) As Pallet
    Dim obj_process As process

    ' process level
    Set obj_process = new_ctrl_process.create_process(obj_record)
    
    ' #implement
      ' check if obj_process status is open or delete
    If Not obj_process.byte_status = new_ctrl_process.BYTE_STATUS_DELETE Then
        Set create_record = New Pallet
        ' pallet level
        If obj_record.str_combi_vhu_to = "" Then
            create_record.str_id = obj_record.str_combi_vhu_from
        Else
            create_record.str_id = obj_record.str_combi_vhu_to
        End If
        create_record.str_material = obj_record.str_material
        create_record.str_material_vendor = obj_record.str_gr_vendor
        create_record.str_material_bin_storage_group = obj_record.str_storage_group_material
        
        'create_record.str_source = obj_record.str_bin_from
        'create_record.str_destination = obj_record.str_machine_transport_ref
        
        'create_record.str_source_destination = obj_record.str_wc_shipto
        create_record.lng_stock_unit = 1
          ' add process to pallet
        create_record.obj_process = obj_process
          ' add pallet into local pallets collection
        On Error GoTo INFO_PALLET_AGGREGATION
        add_pallet create_record
        On Error GoTo 0
    End If
    
    ' clean up
    Set obj_process = Nothing
    Exit Function
INFO_PALLET_AGGREGATION:
    aggregate_pallet get_pallet(create_record.str_id), obj_record
'    Set obj_process.obj_master_version = _
'        new_ctrl_process_master.get_master_version_default(obj_record)
    Resume Next
End Function

Public Function update_record(obj_record As DBHistoryRecord, obj_pallet_record As Pallet, byte_action As Byte)
    'swap_pallet_id obj_pallet_record, obj_record
    new_ctrl_process.update_process obj_record, obj_pallet_record.obj_process, byte_action
End Function

Public Function close_record(obj_record As DBHistoryRecord, obj_pallet_record As Pallet)
    new_ctrl_process.close_process obj_pallet_record.obj_process
    
    If obj_pallet_record.obj_process.byte_status = new_ctrl_process.BYTE_STATUS_CLOSED Then
        'Debug.Print "Writing process to db"
        save_closed_pallet obj_pallet_record
    End If
End Function

Public Function delete_record(obj_record As DBHistoryRecord, obj_pallet_record As Pallet)
    ' #implement
      ' check if transaction action is set to delete
    remove_pallet obj_pallet_record
End Function

Public Function save_closed_pallet(obj_pallet As Pallet)
    save_pallet obj_pallet, obj_mdl_finished, obj_mdl_unfinished
End Function

Public Function save_open_pallets()
    Dim obj_pallet As Pallet
    
    ' clear old records
    'new_mdl_data_process.obj_model.obj_unfinished.clear
    
    For Each obj_pallet In col_pallets
        If obj_pallet.obj_process.byte_status = new_ctrl_process.BYTE_STATUS_OPEN Then
            save_pallet obj_pallet, obj_mdl_finished, obj_mdl_unfinished  'new_mdl_data_process.obj_model.obj_unfinished
        Else
            hndl_log.log db_log.TYPE_WARN, str_module, "save_open_pallets", _
                "After processing all history record there's pallet: " & obj_pallet.str_id & " which is closed, but not saved."
        End If
    Next
End Function

'Private Function save_pallet(obj_pallet As Pallet, obj_mdl_data_process As Object)
Private Function save_pallet(obj_pallet As Pallet, obj_mdl_data_process_finished As MDLDataProcessExcel, Optional obj_mdl_data_process_unfinished As MDLDataProcessExcel)
    Dim obj_db_data_process As DBDataProcess
    Dim obj_process_step As ProcessStep
    Dim str_file_date As String
    
    Set obj_db_data_process = New DBDataProcess
    
    If obj_pallet.str_id = "357020105503043505" Then
        DoEvents
    End If
    
    For Each obj_process_step In obj_pallet.obj_process.obj_steps
        'If obj_process_step.byte_status = obj_mdl_data_process.BYTE_STEP_STATUS Then 'new_mdl_data_process.obj_model.obj_finished.BYTE_STEP_STATUS Then
              ' pallet level
            obj_db_data_process.str_pallet = obj_pallet.str_id
            obj_db_data_process.str_material = obj_pallet.str_material
            obj_db_data_process.str_material_vendor = obj_pallet.str_material_vendor
            obj_db_data_process.str_material_bin_storage_group = obj_pallet.str_material_bin_storage_group
            
            obj_db_data_process.str_source = obj_pallet.obj_process.str_source
            obj_db_data_process.str_source_type = obj_pallet.obj_process.str_source_type
            obj_db_data_process.str_destination = obj_pallet.obj_process.str_destination
            obj_db_data_process.str_destination_type = obj_pallet.obj_process.str_destination_type
            obj_db_data_process.lng_stock_unit = obj_pallet.lng_stock_unit
              ' process level
            obj_db_data_process.str_creation_id = obj_pallet.obj_process.obj_master_version.obj_master.str_process_id
            obj_db_data_process.str_version = obj_pallet.obj_process.obj_master_version.str_id
              ' step level
            obj_db_data_process.str_user = obj_process_step.str_user
            obj_db_data_process.str_date_start = obj_process_step.str_date_start
            obj_db_data_process.str_date_end = obj_process_step.str_date_end
            obj_db_data_process.str_bin_from = obj_process_step.str_bin_from
            obj_db_data_process.str_bin_to = obj_process_step.str_bin_to
            obj_db_data_process.str_transaction_type_start = obj_process_step.str_transaction_type_start
            obj_db_data_process.str_transaction_type_end = obj_process_step.str_transaction_type_end
            'obj_db_data_process.str_order_status = obj_process_step.str_order_status
            obj_db_data_process.byte_process_step_status = obj_process_step.byte_process_status
            obj_db_data_process.byte_process_status = obj_process_step.obj_process.byte_status
                        
            str_file_date = Format(DateValue(obj_data_provider_info.obj_period.str_end) + TimeValue(obj_data_provider_info.obj_period.str_end) - TimeValue("00:00:01"), "d.m.yyyy hh:mm:ss")
                        
            If obj_process_step.byte_status = new_db_process_step.BYTE_CLOSED Then
                obj_mdl_data_process_finished.save_record_dynamic obj_db_data_process, str_file_date
            Else
                'obj_db_data_process.str_date_end = str_file_date
                
                obj_mdl_data_process_finished.save_record_dynamic obj_db_data_process, str_file_date
                obj_mdl_data_process_unfinished.save_record_static obj_db_data_process
            End If
                        
'            If obj_process_step.byte_status = new_db_process_step.BYTE_CLOSED Then
'                obj_mdl_data_process_finished.save_record_dynamic obj_db_data_process
'            Else
'                obj_mdl_data_process_unfinished.save_record_static obj_db_data_process
'            End If
            
        'End If
    Next
    
    If obj_pallet.str_id = "357020105503043505" Then
        DoEvents
    End If
    col_pallets.Remove retrieve_id(obj_pallet.str_id) 'retrieve_id_from_pallet(obj_pallet)
End Function

Public Function add_pallet(obj_pallet As Pallet)
    col_pallets.add obj_pallet, retrieve_id(obj_pallet.str_id) 'retrieve_id_from_pallet(obj_pallet)
End Function

Public Function remove_pallet(obj_pallet As Pallet)
    col_pallets.Remove retrieve_id(obj_pallet.str_id)
End Function

Public Function get_pallet(str_id As String) As Pallet
    Set get_pallet = col_pallets.Item(str_id)
End Function


Private Function aggregate_pallet(obj_pallet As Pallet, obj_record As DBHistoryRecord)
    '# implement
    ' enhance calculation of stock unit based on information from DBHistoryRecord - currently DBHistoryRecord is missing this information
    obj_pallet.lng_stock_unit = obj_pallet.lng_stock_unit + 1
End Function

Private Function swap_pallet_id(obj_pallet As Pallet, obj_record As DBHistoryRecord) As Boolean
    swap_pallet_id = False

    If obj_pallet.str_id <> obj_record.str_combi_vhu_to And obj_record.str_combi_vhu_to <> "" Then
        'obj_pallet.obj_process.obj_actual_step.byte_status = new_db_process_step.BYTE_CLOSED
        update_record obj_record, obj_pallet, new_ctrl_process_master_action.BYTE_UPDATE + new_ctrl_process_master_action.BYTE_CLOSE
        'obj_pallet.obj_process.obj_actual_step.byte_process_status = new_db_process_step.BYTE_PROCESS_STATUS_STOP
        close_record obj_record, obj_pallet
        obj_pallet.str_id = obj_record.str_combi_vhu_to
        On Error GoTo TEMP_ERR
        add_pallet obj_pallet
        On Error GoTo 0
    End If
    Exit Function
TEMP_ERR:
    aggregate_pallet get_pallet(obj_pallet.str_id), obj_record
    swap_pallet_id = True
End Function


'Private Function swap_pallet_id(obj_pallet As Pallet, obj_record As DBHistoryRecord)
'    If obj_pallet.str_id <> obj_record.str_combi_vhu_to And obj_record.str_combi_vhu_to <> "" Then
'        remove_pallet obj_pallet
'        obj_pallet.str_id = obj_record.str_combi_vhu_to
'        On Error GoTo TEMP_ERR
'        add_pallet obj_pallet
'        On Error GoTo 0
'    End If
'    Exit Function
'TEMP_ERR:
'    aggregate_pallet get_pallet(obj_pallet.str_id), obj_record
'    DoEvents
'End Function

Private Function retrieve_id(str_pallet_id As String) As String
    retrieve_id = str_pallet_id
End Function




