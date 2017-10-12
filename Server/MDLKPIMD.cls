VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLKPIMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLKPIMD"

Public str_path As String
Public str_file_name As String
Public str_file_appendix As String

Public BOOL_EXTERNAL_DATA_FILE_VISIBILITY As Boolean

' local data
Public STR_WS_NAME As String
Public STR_DATA_FIRST_CELL As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_PROCESS_ID As Integer
Public INT_OFFSET_PROCESS_VERSION As Integer
Public INT_OFFSET_ABBREV As Integer
Public INT_OFFSET_TYPE As Integer
Public INT_OFFSET_KPI_TARGET As Integer
Public INT_OFFSET_CS_NAME As Integer

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

Public col_kpis As Collection

Public Property Get file_path() As String
    file_path = str_path & str_file_name
End Property

Public Function init()
    STR_DATA_FIRST_CELL = "A2"
    STR_ID_SEPARATOR = "-"
    STR_PROCESS_ID_SEPARATOR = ";"
    'STR_FIRST_ROW_TBL = "A1:J1"
    'STR_FIRST_ROW_DATA = "A2:J2"
    
    'STR_MASTER_MISC_ID = "Miscellaneous"
    
    INT_OFFSET_ID = 0
    INT_OFFSET_PROCESS_ID = 1
    INT_OFFSET_PROCESS_VERSION = 2
    INT_OFFSET_ABBREV = 3
    INT_OFFSET_TYPE = 4
    INT_OFFSET_KPI_TARGET = 5
    INT_OFFSET_CS_NAME = 6
    
    STR_TYPE_ON_TIME = "ON_TIME"
    STR_TYPE_FLOW = "FLOW"
    
    Set col_kpis = New Collection
                            
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

Public Function load_data(Optional col_processes As Collection)
    Dim obj_kpi As Object

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING

    open_data
    ' init
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    
    Do While rg_record.Value <> ""
        DoEvents
                                        
        On Error GoTo INFO_NEW_KPI_SETTING
        Set obj_kpi = col_kpis(rg_record.Offset(0, INT_OFFSET_ID).Value)
        On Error GoTo 0
        
        update obj_kpi, rg_record
        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop

    ' clean up
    Set obj_kpi = Nothing
    close_data
    
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING_FINISHED
    Exit Function
INFO_NEW_KPI_SETTING:
    Set obj_kpi = create(rg_record, col_processes)
    col_kpis.add obj_kpi, obj_kpi.str_id 'retrieve_id_from_config(rg_record)
    Resume Next
ERR_CREATION_PROCESS:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Unexpected error occurred during kpi master data creation."
    Resume Next
End Function

' config
Public Function create(rg_record As Range, col_processes As Collection) As Object
    Set create = create_kpi(rg_record)
    create.str_id = rg_record.Offset(0, INT_OFFSET_ID).Value
    create.str_abbrev = rg_record.Offset(0, INT_OFFSET_ABBREV).Value
    create.str_type = rg_record.Offset(0, INT_OFFSET_TYPE).Value
    create.dbl_target = rg_record.Offset(0, INT_OFFSET_KPI_TARGET).Value
    create.str_name = rg_record.Offset(0, INT_OFFSET_CS_NAME).Value
    
    If Not col_processes Is Nothing Then
        create.obj_process = col_processes(rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value)
    End If
End Function

Private Function create_kpi(rg_record) As Object
    Select Case rg_record.Offset(0, INT_OFFSET_TYPE).Value
        Case STR_TYPE_ON_TIME
            Set create_kpi = New KPIOnTime
        Case STR_TYPE_FLOW
    End Select
End Function

Public Function update(obj_kpi As Object, rg_record As Range)
    If Not obj_kpi.obj_process Is Nothing Then
        obj_kpi.add_process_version obj_kpi.obj_process.col_versions(rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION).Value)
    End If
End Function





