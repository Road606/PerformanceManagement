Attribute VB_Name = "new_ctrl_process_master_action"
Option Explicit

Public Const str_module As String = "new_ctrl_process_master_action"

Public Const STR_CREATE As String = "CREATE"
Public Const STR_UPDATE As String = "UPDATE"
Public Const STR_CLOSE As String = "CLOSE"
Public Const STR_DELETE As String = "DELETE"

Public Const BYTE_CREATE As Byte = 1
Public Const BYTE_UPDATE As Byte = 2
Public Const BYTE_CLOSE As Byte = 4
Public Const BYTE_DELETE As Byte = 8

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

Public col_actions As Collection

Public Property Get file_path() As String
    file_path = str_path & str_file_name
End Property

Public Function init()
    STR_DATA_FIRST_CELL = "A2"
    STR_ID_SEPARATOR = "-"
    STR_PROCESS_ID_SEPARATOR = ";"
    
    'Set col_actions = New Collection
End Function

Public Function load_data()
    Dim obj_action As ProcessMasterAction
    Dim bool_create As Boolean

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING
    
    open_data
    ' set range to cell where data begins
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    
    Do While rg_record.Value <> ""
        DoEvents
                                        
        On Error GoTo WARN_ACTION_ALREADY_EXISTS
        Set obj_action = create(rg_record)
        'col_actions.add obj_action, retrieve_id_from_config(rg_record)
        On Error GoTo 0

        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop

    close_data
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING_FINISHED
    
    ' #implement
    ' load transaction conditions
    new_ctrl_process_master_transac.load_data
    
    Exit Function
WARN_ACTION_ALREADY_EXISTS:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Action: " & rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_ACTION).Value & " is already registered."
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

Public Function create(rg_record As Range) As ProcessMasterAction
    Set create = New ProcessMasterAction
    
    create.str_id = rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_ACTION).Value
    create.obj_master = new_ctrl_process_master.get_master(rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_PROCESS_ID).Value)
End Function

Public Function retrieve_id_from_config(rg_record As Range) As String
    retrieve_id_from_config = _
        retrieve_id(rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_PROCESS_ID).Value, _
            rg_record.Offset(0, new_db_process_master_action.INT_OFFSET_ACTION).Value)
End Function

Public Function retrieve_id(str_process_id As String, str_action_id As String) As String
    retrieve_id = str_process_id & STR_ID_SEPARATOR & str_action_id
End Function








