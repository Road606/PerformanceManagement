VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLAreaMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLAreaMD"

Public str_path As String
Public str_file_name As String
Public str_file_appendix As String

Public BOOL_EXTERNAL_DATA_FILE_VISIBILITY As Boolean

' local data
Public STR_WS_NAME As String
Public STR_DATA_FIRST_CELL As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_CS_NAME As Integer

' objects
Public STR_ID_SEPARATOR As String
Public STR_PROCESS_ID_SEPARATOR As String

' status bar
Public STR_STATUS_BAR_PREFIX As String
Public STR_STATUS_BAR_PREFIX_LOADING As String
Public STR_STATUS_BAR_PREFIX_LOADING_FINISHED As String

Public wb As Workbook

Public col_areas As Collection

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
    INT_OFFSET_CS_NAME = 1
    
    Set col_areas = New Collection
                            
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
    Dim obj_area As Area

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING

    open_data
    ' init
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    
    Do While rg_record.Value <> ""
        DoEvents
                                        
        On Error GoTo ERR_CREATION_PROCESS
        Set obj_area = create(rg_record)
        col_areas.add obj_area, obj_area.str_id 'retrieve_id_from_config(rg_record)
        On Error GoTo 0
        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop

    ' clean up
    Set obj_area = Nothing
    close_data
    
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING_FINISHED
    Exit Function
ERR_CREATION_PROCESS:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Unexpected error occurred during area master data creation."
    Resume Next
End Function

' config
Public Function create(rg_record As Range) As Area
    Set create = New Area
    create.str_id = rg_record.Offset(0, INT_OFFSET_ID).Value
    create.str_name = rg_record.Offset(0, INT_OFFSET_CS_NAME).Value
End Function
  
'Public Function retrieve(str_id As String) As Area
'    Set get_master = col_areas.Item(str_id)
'End Function
  
'Public Function retrieve_id_from_config(rg_record As Range) As String
'    retrieve_id_from_config = _
'        retrieve_id(rg_record.Offset(0, new_db_process_master.INT_OFFSET_CREATION_ID).Value)
'End Function
'
'Private Function retrieve_id(str_process_id As String) As String
'    retrieve_id = str_process_id
'End Function






