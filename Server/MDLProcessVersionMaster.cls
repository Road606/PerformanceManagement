VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLProcessVersionMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLProcessVersionMaster"

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

Public STR_VERSION_DEFAULT As String
Public STR_VERSION_SINGLE As String
Public STR_VERSION_CROSS As String
'
Public STR_CREATION_METHOD_CREATE As String
Public STR_CREATION_METHOD_SUPPLY As String
Public STR_CREATION_METHOD_SUPPLY_HBW As String
'
'Public STR_STEP_ACTION_TYPE_CREATE As String
'Public STR_STEP_ACTION_TYPE_UPDATE As String
'Public STR_STEP_ACTION_TYPE_CLOSE As String

Public wb As Workbook

Public col_versions As Collection

Public Property Get file_path() As String
    file_path = str_path & str_file_name
End Property

Private Sub Class_Initialize()
    STR_DATA_FIRST_CELL = "A2"
    STR_ID_SEPARATOR = "-"
    STR_PROCESS_ID_SEPARATOR = ";"
    'STR_FIRST_ROW_TBL = "A1:J1"
    'STR_FIRST_ROW_DATA = "A2:J2"
    
'    STR_MASTER_MISC_ID = "Miscellaneous"
'
    STR_VERSION_DEFAULT = "0"
    STR_VERSION_SINGLE = "1"
    STR_VERSION_CROSS = "2"
'
    STR_CREATION_METHOD_CREATE = "CREATE"
    STR_CREATION_METHOD_SUPPLY = "CREATE_SUPPLY"
    STR_CREATION_METHOD_SUPPLY_HBW = "CREATE_SUPPLY_HBW"
'
'    STR_STEP_ACTION_TYPE_CREATE = "CREATE"
'    STR_STEP_ACTION_TYPE_UPDATE = "UPDATE"
'    STR_STEP_ACTION_TYPE_CLOSE = "CLOSE"
    
    ' status bar
    STR_STATUS_BAR_PREFIX = "Process Factory->"
    STR_STATUS_BAR_PREFIX_LOADING = "Loading ..."
    STR_STATUS_BAR_PREFIX_LOADING_FINISHED = "Loading has finished"
    
    Set col_versions = New Collection
End Sub

Public Function load_data(col_process_masters As Collection)
    Dim obj_version As ProcessMasterVersion
    Dim bool_create As Boolean

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING

    open_data
    ' init
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    
    Do While rg_record.Value <> ""
        DoEvents
                                        
        On Error GoTo WARN_VERSION_ALREADY_EXISTS
        Set obj_version = create(rg_record, col_process_masters)
        On Error GoTo 0

        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop

    close_data
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING_FINISHED
    
    ' # implement
    'new_ctrl_process_master_step.load_data
    Exit Function
WARN_VERSION_ALREADY_EXISTS:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Version: " & rg_record.Offset(0, new_db_process_master_version.INT_OFFSET_VERSION).Value & " is already registered."
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

Public Function create(rg_record As Range, col_process_masters As Collection) As ProcessMasterVersion
    Set create = New ProcessMasterVersion
    create.str_id = rg_record.Offset(0, new_db_process_master_version.INT_OFFSET_VERSION).Value
    create.str_name = rg_record.Offset(0, new_db_process_master_version.INT_OFFSET_NAME).Value
    
    If col_process_masters.Count > 0 Then
        ' # implement - INT_OFFSET_PROCESS_ID should be loaded from config not from module
        create.obj_master = col_process_masters.Item( _
            rg_record.Offset(0, new_db_process_master_version.INT_OFFSET_PROCESS_ID).Value)
    End If
End Function

' process version
Public Function retrieve(str_master_id As String, str_version_id As String) As ProcessMasterVersion
    Dim obj_master As ProcessMaster
    
    Set obj_master = new_ctrl_process_master.get_master(str_master_id)
    Set retrieve = obj_master.get_version(str_version_id)
End Function


'Public Function get_from_record(obj_record As DBHistoryToProcessRecord) As ProcessMasterVersion
'    Dim obj_master As ProcessMaster
'
'    Set obj_master = new_ctrl_process_master.get_from_record(obj_record)
'    Set get_from_record = _
'        obj_master.col_versions( _
'            resolve_master_version_from_history_to_process_record(obj_record, obj_master))
'End Function

'Public Function get_master_version_default(obj_record As DBHistoryToProcessRecord) As ProcessMasterVersion
'    Dim obj_master As ProcessMaster
'
'    Set obj_master = get_master(STR_MASTER_MISC_ID)
'    Set get_master_version_default = _
'        obj_master.col_versions( _
'            resolve_master_version_from_history_to_process_record(obj_record, obj_master))
'End Function

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

'Public Function retrieve_id_from_config(rg_record As Range) As String
'    retrieve_id_from_config = _
'        retrieve_id( _
'            rg_record.Offset(0, new_db_process_master_version.INT_OFFSET_PROCESS_ID).Value, _
'            rg_record.Offset(0, new_db_process_master_version.INT_OFFSET_VERSION).Value)
'End Function

Public Function retrieve_id(str_master_id As String, str_id As String) As String
    retrieve_id = str_master_id & STR_ID_SEPARATOR & str_id
End Function
