VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLProcessMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLProcessMaster"

Public str_path As String
Public str_file_name As String
Public str_file_appendix As String

Public BOOL_EXTERNAL_DATA_FILE_VISIBILITY As Boolean

' local data
Public STR_WS_NAME As String
Public STR_DATA_FIRST_CELL As String

' objects
Public STR_ID_SEPARATOR As String
Public STR_PROCESS_ID_SEPARATOR As String

' status bar
Public STR_STATUS_BAR_PREFIX As String
Public STR_STATUS_BAR_PREFIX_LOADING As String
Public STR_STATUS_BAR_PREFIX_LOADING_FINISHED As String

Public STR_MASTER_MISC_ID As String

Public wb As Workbook

Public col_process_masters As Collection

Public STR_TRANSACTION_TYPE_ALLOWED_SEPARATOR As String
Public str_transaction_type_allowed As String

Public Property Get file_path() As String
    file_path = str_path & str_file_name
End Property

Public Function init()
    STR_DATA_FIRST_CELL = "A2"
    STR_ID_SEPARATOR = "-"
    STR_PROCESS_ID_SEPARATOR = ";"
    'STR_FIRST_ROW_TBL = "A1:J1"
    'STR_FIRST_ROW_DATA = "A2:J2"
    
    STR_MASTER_MISC_ID = "Miscellaneous"
    
    Set col_process_masters = New Collection
    
    STR_TRANSACTION_TYPE_ALLOWED_SEPARATOR = ";"
    
    ' status bar
    STR_STATUS_BAR_PREFIX = "Process Factory->"
    STR_STATUS_BAR_PREFIX_LOADING = "Loading ..."
    STR_STATUS_BAR_PREFIX_LOADING_FINISHED = "Loading has finished"
    
End Function

Public Function open_data()
    If file_path = "" Then
        Set wb = ThisWorkbook
    Else
        Set wb = util_file.open_wb(file_path, is_visible:=BOOL_EXTERNAL_DATA_FILE_VISIBILITY)
    End If
End Function

Public Function close_data()
    If Not wb Is ThisWorkbook Then
        Windows(wb.Name).Visible = True
        wb.Close SaveChanges:=False
    End If
    
    Set wb = Nothing
End Function

Public Function load_data()
    Dim obj_process_master As ProcessMaster
    Dim bool_create As Boolean

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING

    open_data
    ' init
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    
    Do While rg_record.Value <> ""
        DoEvents
                                        
        On Error GoTo ERR_CREATION_PROCESS
        Set obj_process_master = create(rg_record)
        col_process_masters.add obj_process_master, obj_process_master.str_process_id 'retrieve_id_from_config(rg_record)
        On Error GoTo 0
        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop

    ' clean up
    Set obj_process_master = Nothing
    close_data
    
    'new_ctrl_process_master_action.load_data
    'add_allowed_transaction_type
    
    'new_ctrl_process_master_version.load_data
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING_FINISHED
    Exit Function
ERR_CREATION_PROCESS:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Unexpected error occurred during process master creation."
    Resume Next
End Function

' config
Public Function create(rg_record As Range) As ProcessMaster
    Dim str_creation_parameters As Variant
    Dim str_creation_single As Variant
    Dim int_index As Integer

    Set create = New ProcessMaster
    create.str_process_id = rg_record.Offset(0, new_db_process_master.INT_OFFSET_CREATION_ID).Value
    create.str_type = rg_record.Offset(0, new_db_process_master.INT_OFFSET_TYPE).Value
    create.str_subtype = rg_record.Offset(0, new_db_process_master.INT_OFFSET_SUBTYPE).Value
    create.str_version_determinant = rg_record.Offset(0, new_db_process_master.INT_OFFSET_VERSION_DETERMINANT).Value
End Function
  
Public Function get_master(str_id As String) As ProcessMaster
    Set get_master = col_process_masters.Item(str_id)
End Function
  
Public Function retrieve_id_from_config(rg_record As Range) As String
    retrieve_id_from_config = _
        retrieve_id(rg_record.Offset(0, new_db_process_master.INT_OFFSET_CREATION_ID).Value)
End Function
  
' record
'Public Function get_from_record(obj_record As DBHistoryToProcessRecord) As ProcessMaster
'    Dim obj_master As ProcessMaster
'    Dim obj_action As ProcessMasterAction
'    Dim obj_condition As TransactionCondition
'    Dim bool_is_found As Boolean
'
'    For Each obj_master In col_process_masters
'        Set obj_action = obj_master.col_actions.Item(new_ctrl_process_master_action.STR_CREATE)
'        For Each obj_condition In obj_action.col_conditions
'            bool_is_found = obj_condition.is_match(obj_record)
'            If bool_is_found Then
'                Set get_from_record = obj_master
'                Exit For
'            End If
'        Next
'
'        If bool_is_found Then
'            Exit For
'        End If
'    Next
'End Function
  
'Public Function get_from_record(obj_record As DBHistoryToProcessRecord) As ProcessMaster
'    Set get_from_record = get_master(retrieve_id_from_history_to_process_record(obj_record))
'End Function
'

'
'  ' config

'
'Public Function retrieve_id_from_history_to_process_record(obj_record As DBHistoryToProcessRecord) As String
'    If obj_record.str_process_id = "" Then
'        retrieve_id_from_history_to_process_record = _
'            retrieve_id( _
'                obj_record.str_additional_transaction_type_started & _
'                STR_PROCESS_ID_SEPARATOR & _
'                obj_record.str_additional_transaction_code & _
'                STR_PROCESS_ID_SEPARATOR & _
'                obj_record.str_process_step_user & _
'                STR_PROCESS_ID_SEPARATOR & _
'                obj_record.str_additional_task_list_type)
'    Else
'        retrieve_id_from_history_to_process_record = obj_record.str_process_id
'    End If
'End Function

  'Public Function retrieve_id(str_creation_id As String) As String
Private Function retrieve_id(str_process_id As String) As String
    retrieve_id = str_process_id
End Function

'' process version
'Public Function get_master_version_from_record(obj_record As DBHistoryToProcessRecord) As ProcessMasterVersion
'    Dim obj_master As ProcessMaster
'
'    Set obj_master = get_from_record(obj_record)
'    Set get_master_version_from_record = _
'        obj_master.col_versions( _
'            resolve_master_version_from_history_to_process_record(obj_record, obj_master))
'End Function
'
'Public Function get_master_version_default(obj_record As DBHistoryToProcessRecord) As ProcessMasterVersion
'    Dim obj_master As ProcessMaster
'
'    Set obj_master = get_master(STR_MASTER_MISC_ID)
'    Set get_master_version_default = _
'        obj_master.col_versions( _
'            resolve_master_version_from_history_to_process_record(obj_record, obj_master))
'End Function
'
'Public Function resolve_master_version_from_history_to_process_record(obj_record As DBHistoryToProcessRecord, obj_master As ProcessMaster) As String
'    Dim obj_version_resolver As Object
'
'    Select Case obj_master.str_version_determinant
'        Case new_ctrl_process_master_version.STR_CREATION_METHOD_SUPPLY
'            Set obj_version_resolver = New VersionOutbound
'        Case new_ctrl_process_master_version.STR_CREATION_METHOD_CREATE
'            Set obj_version_resolver = New VersionSingle
'    End Select
'
'    resolve_master_version_from_history_to_process_record = obj_version_resolver.retrieve(obj_record)
'End Function

'Public Function add_allowed_transaction_type()
'    Dim col_allowed_transaction As Collection
'    Dim obj_master As ProcessMaster
'    Dim obj_transaction As TransactionType
'    Dim var_transaction As Variant
'
'    Set col_allowed_transaction = New Collection
'
'    For Each obj_master In col_process_masters
'        For Each obj_transaction In obj_master.col_transactions
'            On Error GoTo INFO_TRANSACTION_EXISTS
'            col_allowed_transaction.add obj_transaction.str_id, obj_transaction.str_id
'            On Error GoTo 0
'        Next
'    Next
'
'    For Each var_transaction In col_allowed_transaction
'        str_transaction_type_allowed = var_transaction & STR_TRANSACTION_TYPE_ALLOWED_SEPARATOR & str_transaction_type_allowed
'    Next
'    Exit Function
'INFO_TRANSACTION_EXISTS:
'    Resume Next
'End Function
'
'Public Function is_allowed_transaction_type(str_transaction_type As String)
'    is_allowed_transaction_type = False
'
'    If InStr(1, str_transaction_type_allowed, str_transaction_type) > 0 Then
'        is_allowed_transaction_type = True
'    End If
'End Function



