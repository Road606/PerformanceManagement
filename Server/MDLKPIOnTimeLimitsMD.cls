VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLKPIOnTimeLimitsMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLKPIOnTimeLimitsMD"

Public str_path As String
Public str_file_name As String
Public str_file_appendix As String

Public BOOL_EXTERNAL_DATA_FILE_VISIBILITY As Boolean

' local data
Public STR_WS_NAME As String
Public STR_DATA_FIRST_CELL As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_KPI_ID As Integer
Public INT_OFFSET_PROCESS_ID As Integer
Public INT_OFFSET_PROCESS_VERSION_ID As Integer
Public INT_OFFSET_CHECKPOINT_PLACE As Integer
Public INT_OFFSET_CHECKPOINT_PLACE_TYPE As Integer
Public INT_OFFSET_CHECKPOINT_PLACE_NEXT As Integer
Public INT_OFFSET_MAT_GRP_WH_ID As Integer
Public INT_OFFSET_SOURCE As Integer
Public INT_OFFSET_LIMIT As Integer

' objects
Public STR_ID_SEPARATOR As String
Public STR_PROCESS_ID_SEPARATOR As String

' status bar
Public STR_STATUS_BAR_PREFIX As String
Public STR_STATUS_BAR_PREFIX_LOADING As String
Public STR_STATUS_BAR_PREFIX_LOADING_FINISHED As String

Public wb As Workbook

Public STR_TYPE_ON_TIME As String
Public STR_TYPE_FLOW As String

Public col_kpi_limits As Collection

Public Property Get file_path() As String
    file_path = str_path & str_file_name
End Property

Private Sub Class_Initialize()
    STR_DATA_FIRST_CELL = "A2"
    STR_ID_SEPARATOR = "-"
    STR_PROCESS_ID_SEPARATOR = ";"
    'STR_FIRST_ROW_TBL = "A1:J1"
    'STR_FIRST_ROW_DATA = "A2:J2"
    
    'STR_MASTER_MISC_ID = "Miscellaneous"
    
    INT_OFFSET_ID = 0
    INT_OFFSET_KPI_ID = 1
    INT_OFFSET_PROCESS_ID = 2
    INT_OFFSET_PROCESS_VERSION_ID = 3
    INT_OFFSET_CHECKPOINT_PLACE = 4
    INT_OFFSET_CHECKPOINT_PLACE_TYPE = 5
    INT_OFFSET_CHECKPOINT_PLACE_NEXT = 6
    INT_OFFSET_MAT_GRP_WH_ID = 7
    INT_OFFSET_SOURCE = 8
    INT_OFFSET_LIMIT = 9
    
    STR_TYPE_ON_TIME = "ON_TIME"
    STR_TYPE_FLOW = "FLOW"
    
    Set col_kpi_limits = New Collection
                            
    ' status bar
    STR_STATUS_BAR_PREFIX = "Process Factory->"
    STR_STATUS_BAR_PREFIX_LOADING = "Loading ..."
    STR_STATUS_BAR_PREFIX_LOADING_FINISHED = "Loading has finished"
    
End Sub

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

Public Function load_data(Optional col_kpis As Collection)
    Dim obj_kpi_limit As Object
    Dim obj_kpi_last_limit As KPIOnTimeLimitCheckpoint

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING

    open_data
    ' init
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    
    Do While rg_record.Value <> ""
        DoEvents
                                        
        'On Error GoTo ERR_CREATION_PROCESS
        create rg_record, col_kpis, obj_kpi_last_limit
        On Error GoTo 0
        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop

    ' clean up
    Set obj_kpi_limit = Nothing
    close_data
    
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING_FINISHED
    Exit Function
ERR_CREATION_PROCESS:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Unexpected error occurred during kpi master data creation."
    Resume Next
End Function

' config
Public Function create(rg_record As Range, col_kpis As Collection, obj_kpi_last_limit As KPIOnTimeLimitCheckpoint) 'As KPIOnTimeLimitProcVersion
    Dim obj_limit_process As KPIOnTimeLimitProc
    Dim obj_limit_process_version As KPIOnTimeLimitProcVersion
    Dim obj_limit_mat_grp_wh As KPIOnTimeLimitMatGrpWH
    Dim obj_limit_source As KPIOnTimeLimitSource
    Dim obj_limit_checkpoint_place As KPIOnTimeLimitCheckpoint
    
    On Error GoTo INFO_NEW_PROCESS
    Set obj_limit_process = col_kpi_limits(rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value)
    On Error GoTo 0
    
    On Error GoTo INFO_NEW_PROCESS_VERSION
    Set obj_limit_process_version = obj_limit_process.col_limits_version(rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION_ID).Value)
    On Error GoTo 0

    On Error GoTo INFO_NEW_MAT_GRP_WH
    Set obj_limit_mat_grp_wh = obj_limit_process_version.col_limits_mat_grp_wh(rg_record.Offset(0, INT_OFFSET_MAT_GRP_WH_ID).Value)
    On Error GoTo 0
    
    On Error GoTo INFO_NEW_SOURCE
    Set obj_limit_source = obj_limit_mat_grp_wh.col_limits_source(rg_record.Offset(0, INT_OFFSET_SOURCE).Value)
    On Error GoTo 0
    
    Set obj_limit_checkpoint_place = New KPIOnTimeLimitCheckpoint
    obj_limit_checkpoint_place.str_id = rg_record.Offset(0, INT_OFFSET_CHECKPOINT_PLACE).Value
    obj_limit_checkpoint_place.str_type = rg_record.Offset(0, INT_OFFSET_CHECKPOINT_PLACE_TYPE).Value
    obj_limit_checkpoint_place.str_next_checkpoint_id = rg_record.Offset(0, INT_OFFSET_CHECKPOINT_PLACE_NEXT).Value
    obj_limit_checkpoint_place.obj_limit = rg_record.Offset(0, INT_OFFSET_LIMIT).Value
    If Not obj_kpi_last_limit Is Nothing Then
        Set obj_limit_checkpoint_place.obj_previous_limit = obj_kpi_last_limit
        Set obj_kpi_last_limit.obj_next_limit = obj_limit_checkpoint_place
    End If
    
    Set obj_kpi_last_limit = obj_limit_checkpoint_place
    obj_limit_source.add_limit obj_limit_checkpoint_place
    
    Exit Function
INFO_NEW_PROCESS:
    Set obj_limit_process = create_process(rg_record, col_kpis)
    Set obj_kpi_last_limit = Nothing
    Resume Next
INFO_NEW_PROCESS_VERSION:
    Set obj_limit_process_version = create_process_version(rg_record, obj_limit_process)
    Set obj_kpi_last_limit = Nothing
    Resume Next
INFO_NEW_MAT_GRP_WH:
    Set obj_limit_mat_grp_wh = create_mat_grp_wh(rg_record, obj_limit_process_version)
    Set obj_kpi_last_limit = Nothing
    Resume Next
INFO_NEW_SOURCE:
    Set obj_limit_source = create_source(rg_record, obj_limit_mat_grp_wh)
    Set obj_kpi_last_limit = Nothing
    Resume Next
End Function

Public Function create_process(rg_record As Range, col_kpis As Collection) As KPIOnTimeLimitProc
    Set create_process = New KPIOnTimeLimitProc
    create_process.str_id = rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value
    
    If Not col_kpis Is Nothing Then
        create_process.obj_kpi = col_kpis(rg_record.Offset(0, INT_OFFSET_KPI_ID).Value)
    End If
    
    col_kpi_limits.add create_process, create_process.str_id
End Function

Public Function create_process_version(rg_record As Range, obj_limit_process As KPIOnTimeLimitProc) As KPIOnTimeLimitProcVersion
    Set create_process_version = New KPIOnTimeLimitProcVersion
    create_process_version.str_id = rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION_ID).Value
    
    obj_limit_process.add_limit create_process_version
End Function

Public Function create_mat_grp_wh(rg_record As Range, obj_limit_process_version As KPIOnTimeLimitProcVersion) As KPIOnTimeLimitMatGrpWH
    Set create_mat_grp_wh = New KPIOnTimeLimitMatGrpWH
    create_mat_grp_wh.str_id = rg_record.Offset(0, INT_OFFSET_MAT_GRP_WH_ID).Value
    
    obj_limit_process_version.add_limit create_mat_grp_wh
End Function

Public Function create_source(rg_record As Range, obj_limit_mat_grp_wh As KPIOnTimeLimitMatGrpWH) As KPIOnTimeLimitSource
    Set create_source = New KPIOnTimeLimitSource
    create_source.str_id = rg_record.Offset(0, INT_OFFSET_SOURCE).Value
    
    obj_limit_mat_grp_wh.add_limit create_source
End Function

'Public Function retrieve_id_by_foreign_keys( _
'        str_process_version_id As String, str_place_from_id As String, _
'        str_place_to_id As String, str_kpi_id As String) As String
'    retrieve_id_by_foreign_keys = _
'        str_process_version_id & STR_ID_SEPARATOR & _
'        str_place_from_id & STR_ID_SEPARATOR & _
'        str_place_to_id & STR_ID_SEPARATOR & _
'        str_kpi_id
'End Function






