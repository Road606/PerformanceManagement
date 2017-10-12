VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLMaterialGroupWHResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Public col_groups As Collection
'
'Public Function retrieve_id_by_source_bin_storage_group(str_source As String, str_bin_storage_group) As String
'    retrieve_id_by_source_bin_storage_group = "default"
'End Function
'
'Private Sub Class_Initialize()
'    Set col_groups = New Collection
'    col_groups.add "default", "default"
'End Sub

Option Explicit

Private Const str_module As String = "MDLMaterialGroupWH"

Public str_path As String
Public str_file_name As String
Public str_file_appendix As String

Public BOOL_EXTERNAL_DATA_FILE_VISIBILITY As Boolean

' local data
Public STR_WS_NAME As String
Public STR_DATA_FIRST_CELL As String

Public INT_OFFSET_MATERIAL_GROUP_WH As Integer
Public INT_OFFSET_VENDOR As Integer
Public INT_OFFSET_BIN_STORAGE_GROUP_PATTERN As Integer
Public INT_OFFSET_CS_NAME As Integer

' objects
Public STR_ID_SEPARATOR As String
Public STR_PROCESS_ID_SEPARATOR As String

' status bar
Public STR_STATUS_BAR_PREFIX As String
Public STR_STATUS_BAR_PREFIX_LOADING As String
Public STR_STATUS_BAR_PREFIX_LOADING_FINISHED As String

Public wb As Workbook

Public col_vendors As Collection

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
    
    INT_OFFSET_MATERIAL_GROUP_WH = 0
    INT_OFFSET_VENDOR = 1
    INT_OFFSET_BIN_STORAGE_GROUP_PATTERN = 2
    INT_OFFSET_CS_NAME = 3

    Set col_vendors = New Collection
                            
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

Public Function load_data()

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING

    open_data
    ' init
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    
    Do While rg_record.Value <> ""
        DoEvents
                                        
        'On Error GoTo ERR_CREATION_MATERIAL_GROUP_WH
        load rg_record
        On Error GoTo 0
        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop

    ' clean up
    close_data
    
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING_FINISHED
    Exit Function
ERR_CREATION_MATERIAL_GROUP_WH:
    hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Unexpected error occurred during kpi master data creation."
    Resume Next
End Function

' config
Public Function load(rg_record As Range)
    Dim obj_vendor As Vendor
    Dim obj_material_group_wh_resolver As MaterialGroupWHResolver
        
    On Error GoTo INFO_NEW_VENDOR
    Set obj_vendor = col_vendors(rg_record.Offset(0, INT_OFFSET_VENDOR).Value)
    On Error GoTo 0
    
    Set obj_material_group_wh_resolver = New MaterialGroupWHResolver
    obj_material_group_wh_resolver.str_material_group_wh_id = rg_record.Offset(0, INT_OFFSET_MATERIAL_GROUP_WH).Value
    obj_material_group_wh_resolver.str_bin_storage_group_pattern = rg_record.Offset(0, INT_OFFSET_BIN_STORAGE_GROUP_PATTERN).Value
    ' add to vendor
    obj_vendor.add_material_group_wh_resolver obj_material_group_wh_resolver
    
    Exit Function
INFO_NEW_VENDOR:
    Set obj_vendor = create_vendor(rg_record)
    Resume Next
End Function

Public Function create_vendor(rg_record As Range) As Vendor
    Set create_vendor = New Vendor
    create_vendor.str_id = rg_record.Offset(0, INT_OFFSET_VENDOR).Value
    
    col_vendors.add create_vendor, create_vendor.str_id
End Function
