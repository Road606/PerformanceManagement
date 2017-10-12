Attribute VB_Name = "new_ctrl_pallet_old"
'Option Explicit
'
'Public Const STR_MODULE As String = "new_ctrl_pallet"
'
'Public col_pallets As Collection
'
'Public Function init()
'    Set col_pallets = New Collection
'End Function
'
'Public Function process_record(obj_record As DBHistoryToProcessRecord)
'    Dim obj_pallet As Pallet
'
'    On Error GoTo INFO_NEW_PALLET
'    Set obj_pallet = get_pallet_from_record(obj_record)
'    On Error GoTo 0
'
'    ' based on process master version check what to do with current transaction
'    ' #implement UPDATE/CLOSE/DELETE
'    update_record obj_record, obj_pallet
'
'    If obj_pallet.obj_process.byte_status = new_ctrl_process.BYTE_STATUS_CLOSED Then
'        Debug.Print "Writing process to db"
'        save_closed_pallet obj_pallet
'    End If
'    Exit Function
'INFO_NEW_PALLET:
'    Set obj_pallet = create_record(obj_record)
'    Resume Next
'End Function
'
'Public Function create_record(obj_record As DBHistoryToProcessRecord) As Pallet
'    Dim obj_process As process
'
'    Set create_record = New Pallet
'    ' pallet level
'    create_record.str_id = obj_record.str_pallet_id
'    create_record.str_material = obj_record.str_pallet_material
'    create_record.str_source_destination = obj_record.str_pallet_source_destination
'    ' process level
'      ' create
'    Set obj_process = New process
'    obj_process.byte_status = new_ctrl_process.BYTE_STATUS_OPEN
'      ' retrieve master for process
'    On Error GoTo INFO_MASTER_NOT_FOUND
'    Set obj_process.obj_master_version = new_ctrl_process_master.get_from_record(obj_record)
'    On Error GoTo 0
'      ' add process to pallet
'    create_record.obj_process = obj_process
'      ' add pallet into local pallets collection
'    add_pallet create_record
'
'    ' clean up
'    Set obj_process = Nothing
'    Exit Function
'INFO_MASTER_NOT_FOUND:
''    Set obj_process.obj_master_version = _
''        new_ctrl_process_master.get_master_version_default(obj_record)
''    Resume Next
'End Function
'
'Public Function update_record(obj_record As DBHistoryToProcessRecord, obj_pallet_record As Pallet)
''    new_ctrl_process.process_step obj_record, obj_pallet_record.obj_process
'End Function
'
'Public Function delete_record(obj_record As DBHistoryToProcessRecord, obj_pallet_record As Pallet)
'    ' #implement
'      ' check if transaction action is set to delete
'End Function
'
'Public Function save_closed_pallet(obj_pallet As Pallet)
'    save_pallet obj_pallet, new_mdl_data_process.obj_model.obj_finished ' #change - obj_model should be direct attribute of this module
'End Function
'
'Public Function save_open_pallets()
'    Dim obj_pallet As Pallet
'
'    ' clear old records
'    new_mdl_data_process.obj_model.obj_unfinished.clear
'
'    For Each obj_pallet In col_pallets
'        If obj_pallet.obj_process.byte_status = new_ctrl_process.BYTE_STATUS_OPEN Then
'            save_pallet obj_pallet, new_mdl_data_process.obj_model.obj_unfinished
'        End If
'    Next
'End Function
'
'Private Function save_pallet(obj_pallet As Pallet, obj_mdl_data_process As Object)
'    Dim obj_db_data_process As DBDataProcess
'    Dim obj_process_step As ProcessStep
'
'    Set obj_db_data_process = New DBDataProcess
'
'    For Each obj_process_step In obj_pallet.obj_process.obj_steps
'        If obj_process_step.byte_status = new_mdl_data_process.obj_model.obj_finished.BYTE_STEP_STATUS Then
'              ' pallet level
'            obj_db_data_process.str_pallet = obj_pallet.str_id
'            obj_db_data_process.str_material = obj_pallet.str_material
'            obj_db_data_process.str_source_destination = obj_pallet.str_source_destination
'              ' process level
'            obj_db_data_process.str_creation_id = obj_pallet.obj_process.obj_master_version.obj_master.str_process_id
'            obj_db_data_process.str_version = obj_pallet.obj_process.obj_master_version.str_id
'              ' step level
'            obj_db_data_process.str_user = obj_process_step.str_user
'            obj_db_data_process.str_date_start = obj_process_step.str_date_start
'            obj_db_data_process.str_date_end = obj_process_step.str_date_end
'            obj_db_data_process.str_bin_from = obj_process_step.str_bin_from
'            obj_db_data_process.str_bin_to = obj_process_step.str_bin_to
'            obj_db_data_process.str_order_status = obj_process_step.str_order_status
'
'            obj_mdl_data_process.save_record obj_db_data_process
'        End If
'    Next
'
'    col_pallets.Remove retrieve_id_from_pallet(obj_pallet)
'End Function
'
'Public Function add_pallet(obj_pallet As Pallet)
'    col_pallets.add obj_pallet, retrieve_id_from_pallet(obj_pallet)
'End Function
'
'' get pallet from local collection
'Public Function get_pallet_from_record(obj_record As DBHistoryToProcessRecord) As Pallet
'    Set get_pallet_from_record = get_pallet( _
'        retrieve_id_from_record(obj_record))
'End Function
'
'Private Function get_pallet(str_id As String) As Pallet
'    Set get_pallet = col_pallets.Item(str_id)
'End Function
'
'' resolve id
'Public Function retrieve_id_from_pallet(obj_pallet As Pallet) As String
'    retrieve_id_from_pallet = retrieve_id(obj_pallet.str_id)
'End Function
'
'Public Function retrieve_id_from_record(obj_record As DBHistoryToProcessRecord) As String
'    If obj_record.str_pallet_id = "" Then ' str_pallet id is set directly record origins from data.process database
'        obj_record.str_pallet_id = obj_record.str_additional_vhu_to
'
'        If obj_record.str_pallet_id = "" Then
'            obj_record.str_pallet_id = obj_record.str_additional_vhu_from
'        End If
'    End If
'
'    retrieve_id_from_record = retrieve_id(obj_record.str_pallet_id)
'End Function
'
'Private Function retrieve_id(str_pallet_id As String) As String
'    retrieve_id = str_pallet_id
'End Function
