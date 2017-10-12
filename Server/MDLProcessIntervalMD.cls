VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLProcessIntervalMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "MDLProcessIntervalMD"

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_PROCESS_ID As Integer
Public INT_OFFSET_LOWER_LIMIT As Integer
Public INT_OFFSET_UPPER_LIMIT As Integer
Public INT_OFFSET_MINUTE As Integer
Public INT_OFFSET_HOUR_OFFSET As Integer
Public INT_OFFSET_IS_COMPLETE_OFFSET As Integer

Public obj_data_provider As FileExcelDataProvider
Public col_process_intervals As Collection
Public col_processes As Collection

Public Function load()
    'Set obj_data_provider = New FileExcelDataProvider
    'obj_data_provider.str_ws_name = str_ws_name
    'obj_data_provider.add_listener Me
    obj_data_provider.load_data
End Function

' # interface method for listening FileExcelDataProvider
Public Function load_record(rg_record As Range)
    Dim obj_interval As Object
    
    Set obj_interval = create(rg_record, col_processes)
    col_process_intervals.add obj_interval, obj_interval.str_id
End Function

Public Function create(rg_record As Range, col_processes As Collection) As ProcessInterval
    Set create = New ProcessInterval
    create.str_id = rg_record.Offset(0, INT_OFFSET_ID).Value
    create.int_lower_limit = rg_record.Offset(0, INT_OFFSET_LOWER_LIMIT).Value
    create.int_upper_limit = rg_record.Offset(0, INT_OFFSET_UPPER_LIMIT).Value
    create.int_minute = rg_record.Offset(0, INT_OFFSET_MINUTE).Value
    create.int_hour_offset = rg_record.Offset(0, INT_OFFSET_HOUR_OFFSET).Value
    create.int_is_complete = rg_record.Offset(0, INT_OFFSET_IS_COMPLETE_OFFSET).Value
    
    If Not col_processes Is Nothing Then
        create.obj_process = col_processes(rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value)
    End If
End Function

Private Sub Class_Initialize()
    Set col_process_intervals = New Collection
    
    INT_OFFSET_ID = 0
    INT_OFFSET_PROCESS_ID = 1
    INT_OFFSET_LOWER_LIMIT = 2
    INT_OFFSET_UPPER_LIMIT = 3
    INT_OFFSET_MINUTE = 4
    INT_OFFSET_HOUR_OFFSET = 5
    INT_OFFSET_IS_COMPLETE_OFFSET = 6
End Sub

