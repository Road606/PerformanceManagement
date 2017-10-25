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

Public str_module As String
Public str_process_collection As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_PROCESS_ID As Integer
Public INT_OFFSET_LOWER_LIMIT As Integer
Public INT_OFFSET_UPPER_LIMIT As Integer
Public INT_OFFSET_MINUTE As Integer
Public INT_OFFSET_HOUR_OFFSET As Integer
Public INT_OFFSET_IS_COMPLETE_OFFSET As Integer

Public single_data_provider As FileExcelDataProvider

Public col_listeners As Collection


Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function

Public Function load_data()
    single_data_provider.add_listener Me
    single_data_provider.load_data
End Function

Public Function load_record(rg_record As Range)
    Dim obj_listener As Object
    Dim obj_interval As ProcessInterval
    
    For Each obj_listener In col_listeners
        Set obj_interval = New ProcessInterval
        obj_interval.str_id = rg_record.Offset(0, INT_OFFSET_ID).Value
        obj_interval.int_lower_limit = rg_record.Offset(0, INT_OFFSET_LOWER_LIMIT).Value
        obj_interval.int_upper_limit = rg_record.Offset(0, INT_OFFSET_UPPER_LIMIT).Value
        obj_interval.int_minute = rg_record.Offset(0, INT_OFFSET_MINUTE).Value
        obj_interval.int_hour_offset = rg_record.Offset(0, INT_OFFSET_HOUR_OFFSET).Value
        obj_interval.int_is_complete = rg_record.Offset(0, INT_OFFSET_IS_COMPLETE_OFFSET).Value
    
        If obj_listener.get_collection(str_process_collection).Count > 0 Then
            obj_interval.obj_process = obj_listener.get_collection(str_process_collection)(rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value)
        End If
        obj_listener.get_collection(str_module).add obj_interval, obj_interval.str_id
    Next obj_listener
End Function

Private Sub Class_Initialize()

str_module = new_const_mdl_id.STR_MDL_PROCESS_INTERVAL
str_process_collection = new_const_mdl_id.STR_MDL_PROCESS_MASTER

INT_OFFSET_ID = 0
INT_OFFSET_PROCESS_ID = 1
INT_OFFSET_LOWER_LIMIT = 2
INT_OFFSET_UPPER_LIMIT = 3
INT_OFFSET_MINUTE = 4
INT_OFFSET_HOUR_OFFSET = 5
INT_OFFSET_IS_COMPLETE_OFFSET = 6

Set col_listeners = New Collection

Set single_data_provider = New FileExcelDataProvider

End Sub
