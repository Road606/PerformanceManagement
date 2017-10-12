Attribute VB_Name = "new_ctrl_process_master_step"
Option Explicit

Public Const str_module As String = "new_ctrl_process_master_step"

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

Public Property Get file_path() As String
    file_path = str_path & str_file_name
End Property

Public Function init()
    STR_DATA_FIRST_CELL = "A2"
    STR_ID_SEPARATOR = "-"
    STR_PROCESS_ID_SEPARATOR = ";"
    
End Function

Public Function load_data()
    Dim obj_step As ProcessMasterStep
    Dim col_steps As Collection
    Dim bool_create As Boolean

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING

    open_data
    ' init
    Set col_steps = New Collection
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)

    Do While rg_record.Value <> ""
        DoEvents

        On Error GoTo INFO_NEW_STEP
        Set obj_step = col_steps.Item(retrieve_id_from_config(rg_record))
        On Error GoTo 0

        update obj_step, rg_record

        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop

    Set obj_step = Nothing
    Set col_steps = Nothing
    close_data
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING_FINISHED

    Exit Function
INFO_NEW_STEP:
    On Error GoTo WARN_STEP_ALREADY_EXISTS
    Set obj_step = create(rg_record)
    On Error GoTo 0
    Resume Next
WARN_STEP_ALREADY_EXISTS:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Step defined on row: " & rg_record.Row & " is already registered."
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

Public Function create(rg_record As Range) As ProcessMasterStep
    Dim obj_master As ProcessMaster

    Set create = New ProcessMasterStep
    create.str_place_from = rg_record.Offset(0, new_db_process_master_step.INT_OFFSET_PLACE_FROM).Value
    create.str_place_to = rg_record.Offset(0, new_db_process_master_step.INT_OFFSET_PLACE_TO).Value

    Set obj_master = new_ctrl_process_master.get_master(rg_record.Offset(0, new_db_process_master_step.INT_OFFSET_PROCESS_ID).Value)
    create.obj_version = obj_master.get_version(rg_record.Offset(0, new_db_process_master_step.INT_OFFSET_VERSION).Value)
End Function

Public Function update(obj_process_master_step As ProcessMasterStep, rg_record As Range)
    On Error GoTo WARN_ORDER_ALREADY_EXISTS
    obj_process_master_step.add_order rg_record.Offset(0, new_db_process_master_step.INT_OFFSET_STEP_ORDER).Value
    On Error GoTo 0
    Exit Function
WARN_ORDER_ALREADY_EXISTS:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Step order defined on row: " & rg_record.Row & " is already registered."
    Resume Next
End Function

Public Function retrieve_id(str_process_id As String, str_version_id As String, _
        str_place_from As String, str_place_to As String) As String

    retrieve_id = str_process_id & STR_PROCESS_ID_SEPARATOR & _
        str_version_id & STR_PROCESS_ID_SEPARATOR & _
        str_place_from & STR_PROCESS_ID_SEPARATOR & _
        str_place_to & STR_PROCESS_ID_SEPARATOR
End Function

Public Function retrieve_id_from_config(rg_record As Range) As String
    retrieve_id_from_config = _
        retrieve_id( _
            rg_record.Offset(0, new_db_process_master_step.INT_OFFSET_PROCESS_ID).Value, _
            rg_record.Offset(0, new_db_process_master_step.INT_OFFSET_VERSION).Value, _
            rg_record.Offset(0, new_db_process_master_step.INT_OFFSET_PLACE_FROM).Value, _
            rg_record.Offset(0, new_db_process_master_step.INT_OFFSET_PLACE_TO).Value)
End Function







