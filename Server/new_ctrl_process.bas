Attribute VB_Name = "new_ctrl_process"
Option Explicit

Public Const BYTE_STATUS_DELETE As Byte = 0
Public Const BYTE_STATUS_OPEN As Byte = 1
Public Const BYTE_STATUS_CLOSED As Byte = 2

Public Function create_process(obj_record As DBHistoryRecord) As process
    Dim obj_master As ProcessMaster
    Dim obj_version_resolver As Object

    Set create_process = New process
    
    ' retrieve master for process
    On Error GoTo INFO_MASTER_VERSION_NOT_FOUND
    Set obj_master = get_master(obj_record)
    Set obj_version_resolver = resolve_master_version(obj_record, obj_master)
    create_process.str_source = obj_version_resolver.str_source
    create_process.str_source_type = obj_version_resolver.str_source_type
    create_process.str_destination = obj_version_resolver.str_destination
    create_process.str_destination_type = obj_version_resolver.str_destination_type
    
    Set create_process.obj_master_version = _
        new_ctrl_process_master_version.retrieve(obj_master.str_process_id, obj_version_resolver.str_version_id)
    create_process.byte_status = BYTE_STATUS_OPEN
    On Error GoTo 0
    
    ' # implement
      ' create step
    'new_ctrl_process_step.process_step obj_record, create_process
    Exit Function
INFO_MASTER_VERSION_NOT_FOUND:
    create_process.byte_status = BYTE_STATUS_DELETE
End Function

Public Function update_process(obj_record As DBHistoryRecord, obj_process As process, byte_action As Byte)
    new_ctrl_process_step.process_step obj_record, obj_process, byte_action 'new_ctrl_process_master_action.STR_UDATE
    'obj_process.obj_actual_step.byte_process_status = new_db_process_step.BYTE_PROCESS_STATUS_CONTINUE
End Function

Public Function close_process(obj_process As process)
    'obj_process.obj_actual_step.byte_status = new_db_process_step.BYTE_CLOSED
    'obj_process.obj_actual_step.byte_process_status = new_db_process_step.BYTE_PROCESS_STATUS_STOP
    If obj_process.obj_actual_step.byte_process_status = new_db_process_step.BYTE_PROCESS_STATUS_STOP Then
        obj_process.byte_status = BYTE_STATUS_CLOSED
    End If
End Function

Public Function is_transaciton_valid_for_action(str_action As String, obj_record As DBHistoryRecord, obj_process As process) As Boolean
    On Error GoTo INFO_ACTION_NOT_FOUND
    is_transaciton_valid_for_action = obj_process.obj_master_version.obj_master.get_action(str_action).is_valid_transaction(obj_record)
    On Error GoTo 0
    Exit Function
INFO_ACTION_NOT_FOUND:
    is_transaciton_valid_for_action = False
End Function

Private Function get_master(obj_record As DBHistoryRecord) As ProcessMaster
    Dim obj_current_master As ProcessMaster
    Dim obj_action As ProcessMasterAction
    Dim obj_condition As TransactionCondition
    Dim bool_is_found As Boolean
    
    For Each obj_current_master In new_ctrl_process_master.col_process_masters
        Set obj_action = obj_current_master.col_actions.Item(new_ctrl_process_master_action.STR_CREATE)
        For Each obj_condition In obj_action.col_conditions
            bool_is_found = obj_condition.is_match(obj_record)
            If bool_is_found Then
                Set get_master = obj_current_master
                Exit For
            End If
        Next
        
        If bool_is_found Then
            Exit For
        End If
    Next
End Function

Private Function resolve_master_version(obj_record As DBHistoryRecord, obj_master As ProcessMaster) As Object
    Dim obj_version_resolver As Object
    
    Select Case obj_master.str_version_determinant
        Case new_ctrl_process_master_version.STR_CREATION_METHOD_SUPPLY
            Set resolve_master_version = New VersionOutbound
        Case new_ctrl_process_master_version.STR_CREATION_METHOD_SUPPLY_HBW
            Set resolve_master_version = New VersionOutboundHBW
        Case new_ctrl_process_master_version.STR_CREATION_METHOD_CREATE
            Set resolve_master_version = New VersionSingle
    End Select
    
    resolve_master_version.init obj_record
End Function

'Private Function resolve_master_version(obj_record As DBHistoryRecord, obj_master As ProcessMaster) As String
'    Dim obj_version_resolver As Object
'
'    Select Case obj_master.str_version_determinant
'        Case new_ctrl_process_master_version.STR_CREATION_METHOD_SUPPLY
'            Set obj_version_resolver = New VersionOutbound
'        Case new_ctrl_process_master_version.STR_CREATION_METHOD_SUPPLY_HBW
'            Set obj_version_resolver = New VersionOutboundHBW
'        Case new_ctrl_process_master_version.STR_CREATION_METHOD_CREATE
'            Set obj_version_resolver = New VersionSingle
'    End Select
'
'    resolve_master_version = obj_version_resolver.retrieve(obj_record)
'End Function
