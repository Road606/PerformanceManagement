Attribute VB_Name = "new_ctrl_process_step"
Option Explicit

'Public Const BYTE_STATUS_OPEN As Byte = 1
'Public Const BYTE_STATUS_CLOSED As Byte = 2

Public Function process_step(obj_record As DBHistoryRecord, obj_process As process, byte_action As Byte) 'str_action As String)
    Dim bool_is_single As Boolean
    Dim bool_is_final As Boolean

    ' create
    If Not obj_process.has_open_step() Then
        create_step obj_process
    End If

    ' update
    update_step obj_record, obj_process

    ' close
    If obj_process.obj_actual_step.byte_status = new_db_process_step.BYTE_CLOSED Then
        close_step obj_process, byte_action
    End If
End Function

Public Function create_step(obj_process As process)
    Dim obj_step As ProcessStep

    Set obj_step = New ProcessStep
    obj_step.byte_status = new_db_process_step.BYTE_NEW
    obj_step.byte_process_status = new_db_process_step.BYTE_PROCESS_STATUS_CONTINUE
    obj_process.add_step obj_step
    Set obj_process.obj_actual_step = obj_step
End Function

Public Function create_following_step(obj_process As process)
    Dim obj_actual_step As ProcessStep
    Dim obj_new_step As ProcessStep

    Set obj_actual_step = obj_process.obj_actual_step
    create_step obj_process
    Set obj_new_step = obj_process.obj_actual_step

    obj_new_step.str_date_start = obj_actual_step.str_date_end
    obj_new_step.str_bin_from = obj_actual_step.str_bin_to
    obj_new_step.str_place_from = obj_actual_step.str_place_to
    obj_new_step.str_transaction_type_start = obj_actual_step.str_transaction_type_end
    obj_new_step.str_user = obj_actual_step.str_user

    ' add information about transaction
    ' # implement
    obj_new_step.byte_status = new_db_process_step.BYTE_OPEN
End Function

Public Function update_step(obj_record As DBHistoryRecord, obj_process As process)
    Dim obj_process_step As ProcessStep

    Set obj_process_step = obj_process.obj_actual_step

    If obj_record.is_single_transaction Then
        If obj_process_step.byte_status = new_db_process_step.BYTE_NEW Then
            update_part_one obj_record, obj_process_step
        Else
            update_part_two obj_record, obj_process_step
        End If
    Else
        If obj_process_step.byte_status = new_db_process_step.BYTE_NEW Then
            update_part_one obj_record, obj_process_step
        End If

        update_part_two obj_record, obj_process_step
    End If
End Function

Public Function close_step(obj_process As process, byte_action As Byte)
    ' check if it is last step in process
    If Not byte_action And new_ctrl_process_master_action.BYTE_CLOSE Then
        obj_process.obj_actual_step.byte_process_status = new_db_process_step.BYTE_PROCESS_STATUS_CONTINUE
        create_following_step obj_process
    Else
        If Not obj_process.obj_actual_step.byte_status = new_db_process_step.BYTE_CLOSED Then
            obj_process.obj_actual_step.byte_process_status = new_db_process_step.BYTE_PROCESS_STATUS_CONTINUE
        Else
            obj_process.obj_actual_step.byte_process_status = new_db_process_step.BYTE_PROCESS_STATUS_STOP
        End If
    End If
End Function

'Public Function update_part_one(obj_record As DBHistoryRecord, obj_step As ProcessStep, bool_is_single As Boolean)
Public Function update_part_one(obj_record As DBHistoryRecord, obj_step As ProcessStep)
    obj_step.str_user = obj_record.str_user_name

    obj_step.str_date_start = obj_record.str_transaction_started 'obj_record.str_transaction_started
    obj_step.str_bin_from = obj_record.str_bin_from 'obj_record.str_bin_from
    obj_step.str_place_from = bin_place_grp.get_place_grp(obj_record.str_bin_from)

    obj_step.str_transaction_type_start = obj_record.str_transaction_type_started
    obj_step.byte_status = new_db_process_step.BYTE_OPEN
End Function

'Public Function update_part_two(obj_record As DBHistoryRecord, obj_step As ProcessStep, bool_is_single As Boolean)
Public Function update_part_two(obj_record As DBHistoryRecord, obj_step As ProcessStep)
    obj_step.str_user = obj_record.str_user_name
    If obj_record.str_bin_to <> "" Then
        obj_step.str_bin_to = obj_record.str_bin_to
        obj_step.str_place_to = bin_place_grp.get_place_grp(obj_step.str_bin_to)
    End If
    
    If obj_record.is_single_transaction Then
        obj_step.str_date_end = obj_record.str_transaction_started
        obj_step.str_transaction_type_end = obj_record.str_transaction_type_started
    Else
        obj_step.str_date_end = obj_record.str_transaction_finished
        obj_step.str_transaction_type_end = obj_record.str_transaction_type_finished
    End If

    obj_step.byte_status = new_db_process_step.BYTE_CLOSED
End Function

Private Function evaluate_step_status(obj_process As process)
'    Dim obj_master_step As ProcessMasterStep
'    Dim obj_actual_step As ProcessStep
'
'    Set obj_actual_step = obj_process.obj_actual_step
'    Set obj_master_step = obj_actual_step.get_master()
'
'    If obj_actual_step.str_order_status = "" Then
'        If obj_master_step Is Nothing Then
'            obj_actual_step.str_order_status = new_db_process_step.STR_ORDER_STATUS_NOT_FOUND
'        Else
'            If obj_process.int_last_step_id >= obj_master_step.int_order Then
'                obj_actual_step.str_order_status = new_db_process_step.STR_ORDER_STATUS_BACK
'            ElseIf obj_process.int_last_step_id + 1 = obj_master_step.int_order Then
'                obj_actual_step.str_order_status = new_db_process_step.STR_ORDER_STATUS_OK
'                obj_process.int_last_step_id = obj_master_step.int_order
'            Else
'                obj_actual_step.str_order_status = new_db_process_step.STR_ORDER_STATUS_AHEAD
'                obj_process.int_last_step_id = obj_master_step.int_order
'            End If
'        End If
'    End If
End Function


