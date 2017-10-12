Attribute VB_Name = "new_ctrl_process_master_transac"
Option Explicit

Public Const str_module As String = "new_ctrl_process_master_transac"

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
Public Const STR_STATUS_BAR_PREFIX As String = "Process Factory->"
Public Const STR_STATUS_BAR_PREFIX_LOADING As String = "Loading ..."
Public Const STR_STATUS_BAR_PREFIX_LOADING_FINISHED As String = "Loading has finished"

Public wb As Workbook

Public col_conditions As Collection

Public Property Get file_path() As String
    file_path = str_path & str_file_name
End Property

Public Function init()
    STR_DATA_FIRST_CELL = "A2"
    STR_ID_SEPARATOR = "-"
    STR_PROCESS_ID_SEPARATOR = ";"

    'Set col_conditions = New Collection
End Function

Public Function load_data()
    'Dim obj_condition As ProcessMasterAction
    'Dim obj_condition As TransactionCondition
    Dim bool_create As Boolean

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING

    open_data
    ' init
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)

    Do While rg_record.Value <> ""
        DoEvents

        On Error GoTo WARN_CONDITION_ALREADY_EXISTS
        create rg_record ' condition is automatically assigned to process master action in create method
        On Error GoTo 0

        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop

    close_data
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING_FINISHED
    Exit Function
WARN_CONDITION_ALREADY_EXISTS:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "During creation of condition: " & rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_ACTION).Value & " occured an unxpected error."
    Resume Next
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

Public Function create(rg_record As Range) 'As TransactionCondition
    Dim obj_process_master As ProcessMaster

    Set create = New TransactionCondition

    create.str_type_start = rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_TRANSACTION_TYPE_START).Value
    create.str_type_end = rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_TRANSACTION_TYPE_END).Value
    create.str_code = rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_TRANSACTION_CODE).Value
    create.str_user = rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_USER).Value
    create.str_task_list_type = rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_TASK_LIST_TYPE).Value
    create.str_place_from = rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_PLACE_FROM).Value
    create.str_place_to = rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_PLACE_TO).Value
    
    Set obj_process_master = new_ctrl_process_master.get_master(rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_PROCESS_ID).Value)
    create.obj_action = obj_process_master.get_action(rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_ACTION).Value)
End Function
'
'Public Function update(obj_action As TransactionType, rg_record As Range)
'    Dim obj_condition As TransactionCondition
'    Dim arr_actions As Variant
'    Dim int_index As Integer
'
'    Set obj_condition = New TransactionCondition
'    obj_condition.obj_type = obj_action
'    obj_condition.str_code = rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_TRANSACTION_CODE).Value
'    obj_condition.str_user = rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_USER).Value
'    obj_condition.str_task_list_type = rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_TASK_LIST_TYPE).Value
'    obj_condition.str_place_to = rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_PLACE_TO).Value
'
'    arr_actions = Split(rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_ACTION).Value, STR_CONFIG_ACTION_SEPARATOR)
'
'    For int_index = 0 To UBound(arr_actions)
'        obj_condition.add_action CStr(arr_actions(int_index))
'    Next int_index
'
'    Set obj_condition = Nothing
'End Function

Public Function retrieve_id_from_config(rg_record As Range) As String
    retrieve_id_from_config = _
        retrieve_id(rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_PROCESS_ID).Value, _
            rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_ACTION).Value, _
            rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_TRANSACTION_TYPE_START).Value, _
            rg_record.Offset(0, new_db_process_master_transact.INT_OFFSET_TRANSACTION_TYPE_END).Value)
End Function

Public Function retrieve_id(str_process_id As String, str_action_id As String, str_type_id_start As String, str_type_id_end) As String
    retrieve_id = str_process_id & STR_ID_SEPARATOR & str_action_id & STR_ID_SEPARATOR & str_type_id_start & STR_ID_SEPARATOR & str_type_id_end
End Function










