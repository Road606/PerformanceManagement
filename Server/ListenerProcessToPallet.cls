VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListenerProcessToPallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STR_MODULE_NAME As String = "ListenerProcessToPallet" ' # implement - delete and use only CLASS_NAME
Private Const STR_CLASS_NAME As String = "ListenerProcessToPallet"

Public Function process_record(obj_record As DBDataProcess)
    Dim obj_pallet As Pallet

    ' retrieve pallet from record
    On Error GoTo INFO_NEW_PALLET
    Set obj_pallet = new_ctrl_pallet.get_pallet(obj_record.str_pallet)
    On Error GoTo 0
        
    ' update pallet
    update_pallet obj_pallet, obj_record
    Exit Function
INFO_NEW_PALLET:
    Set obj_pallet = create_pallet(obj_record)
    Resume Next
End Function

'Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
'    ' # no use yet. Just ignore it
'End Function

Private Function create_pallet(obj_record As DBDataProcess) As Pallet
    Dim obj_process As process
    
    Set create_pallet = New Pallet
    
    ' pallet information
    create_pallet.str_id = obj_record.str_pallet
    create_pallet.str_material = obj_record.str_material
    create_pallet.str_material_vendor = obj_record.str_material_vendor
    create_pallet.str_material_bin_storage_group = obj_record.str_material_bin_storage_group
    'create_pallet.str_source_destination = obj_record.str_destination ' # implement - modify according to new properties like source, destination etc.
    On Error GoTo WARN_INVALID_VALUE_STOCK_UNIT
    create_pallet.lng_stock_unit = obj_record.lng_stock_unit
    On Error GoTo 0
    ' process information
    Set obj_process = New process
    obj_process.byte_status = new_ctrl_process.BYTE_STATUS_OPEN
    Set obj_process.obj_master_version = _
        new_ctrl_process_master_version.retrieve(obj_record.str_creation_id, obj_record.str_version)
    obj_process.str_source = obj_record.str_source
    obj_process.str_source_type = obj_record.str_source_type
    obj_process.str_destination = obj_record.str_destination
    obj_process.str_destination_type = obj_record.str_destination_type
    create_pallet.obj_process = obj_process
    ' add pallet
    new_ctrl_pallet.add_pallet create_pallet
    Exit Function
WARN_INVALID_VALUE_STOCK_UNIT:
    hndl_log.log db_log.TYPE_WARN, STR_MODULE_NAME, "create_pallet", "Invalid value for Stock unit on pallet: " & obj_record.str_pallet
End Function

Private Function update_pallet(obj_pallet As Pallet, obj_record As DBDataProcess)
    Dim obj_process As process
    Dim obj_step As ProcessStep
    
    Set obj_process = obj_pallet.obj_process
    Set obj_step = create_step(obj_record)
    obj_process.add_step obj_step
    Set obj_process.obj_actual_step = obj_step
    
'    If obj_step.byte_status = new_db_process_step.BYTE_CLOSED Then
'        new_ctrl_process_step.create_following_step obj_process
'    End If
End Function

Private Function create_step(obj_record As DBDataProcess) As ProcessStep
    Set create_step = New ProcessStep
    
    create_step.str_bin_from = obj_record.str_bin_from
    create_step.str_bin_to = obj_record.str_bin_to
    create_step.str_place_from = bin_place_grp.get_place_grp(obj_record.str_bin_from)
    If obj_record.str_bin_to <> "" Then
        create_step.str_place_to = bin_place_grp.get_place_grp(obj_record.str_bin_to)
    End If
    
    create_step.str_date_start = obj_record.str_date_start
    create_step.str_date_end = obj_record.str_date_end
    
    create_step.str_transaction_type_start = obj_record.str_transaction_type_start
    create_step.str_transaction_type_end = obj_record.str_transaction_type_end
    
    create_step.str_user = obj_record.str_user
    
    create_step.byte_process_status = obj_record.byte_process_step_status
    create_step.byte_status = obj_record.byte_process_status 'obj_record.str_order_status
    
    If create_step.str_date_end = "" Then
        create_step.byte_status = new_db_process_step.BYTE_OPEN
    Else
        create_step.byte_status = new_db_process_step.BYTE_CLOSED
    End If
End Function

Public Function loading_data_is_about_to_start(obj_provider_info As FileExcelDataProviderInfo)
    
End Function

Public Function loading_data_has_finished(obj_provider_info As FileExcelDataProviderInfo)
    
End Function

