VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLAreaStepMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLAreaStepMD"

Public str_path As String
Public str_file_name As String
Public str_file_appendix As String

Public BOOL_EXTERNAL_DATA_FILE_VISIBILITY As Boolean

' local data
Public STR_WS_NAME As String
Public STR_DATA_FIRST_CELL As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_AREA_ID As Integer
Public INT_OFFSET_PLACE_FROM As Integer
Public INT_OFFSET_PLACE_TO As Integer

' objects
Public STR_ID_SEPARATOR As String
Public STR_PROCESS_ID_SEPARATOR As String

' status bar
Public STR_STATUS_BAR_PREFIX As String
Public STR_STATUS_BAR_PREFIX_LOADING As String
Public STR_STATUS_BAR_PREFIX_LOADING_FINISHED As String

Public wb As Workbook

Public col_area_steps As Collection

Public Property Get file_path() As String
    file_path = str_path & str_file_name
End Property

Public Function init()
    STR_DATA_FIRST_CELL = "A2"
    STR_ID_SEPARATOR = "-"
    STR_PROCESS_ID_SEPARATOR = ";"
    
    INT_OFFSET_ID = 0
    INT_OFFSET_AREA_ID = 1
    INT_OFFSET_PLACE_FROM = 2
    INT_OFFSET_PLACE_TO = 3
    
    Set col_area_steps = New Collection
                            
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

Public Function load_data(col_areas As Collection)
    Dim obj_area_step As AreaStep

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING

    open_data
    ' init
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    
    Do While rg_record.Value <> ""
        DoEvents
                                        
        'On Error GoTo ERR_CREATION_PROCESS
        Set obj_area_step = create(rg_record, col_areas)
        col_area_steps.add obj_area_step, obj_area_step.str_id
        On Error GoTo 0
        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop

    ' clean up
    Set obj_area_step = Nothing
    close_data
    
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING_FINISHED
    Exit Function
ERR_CREATION_PROCESS:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Unexpected error occurred during area master data creation."
    Resume Next
End Function

' config
Public Function create(rg_record As Range, col_areas As Collection) As AreaStep
    Set create = New AreaStep
    create.str_id = rg_record.Offset(0, INT_OFFSET_ID).Value
    create.str_place_from = rg_record.Offset(0, INT_OFFSET_PLACE_FROM).Value
    create.str_place_to = rg_record.Offset(0, INT_OFFSET_PLACE_TO).Value
    
    If col_areas.Count > 0 Then
        create.obj_area = col_areas(rg_record.Offset(0, INT_OFFSET_AREA_ID).Value)
    End If
End Function
'
'Public Function retrieve(str_id As String) As Area
'    Set get_master = col_area_steps.Item(str_id)
'End Function








