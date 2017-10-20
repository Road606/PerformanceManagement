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

Public str_module As String
Public areas_collection As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_AREA_ID As Integer
Public INT_OFFSET_PLACE_FROM As Integer
Public INT_OFFSET_PLACE_TO As Integer

Public single_data_provider As FileExcelDataProvider

Public col_listeners As Collection
Public col_areas As Collection

Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function

Public Function load_data()
    single_data_provider.add_listener Me
    single_data_provider.load_data
End Function

Public Function load_record(rg_record As Range)
Dim obj_listener As Object
Dim obj_area_step As AreaStep

For Each obj_listener In Me.col_listeners
    Set obj_area_step = New AreaStep
    obj_area_step.str_id = rg_record.Offset(0, INT_OFFSET_ID).Value
    obj_area_step.str_place_from = rg_record.Offset(0, INT_OFFSET_PLACE_FROM).Value
    obj_area_step.str_place_to = rg_record.Offset(0, INT_OFFSET_PLACE_TO).Value

    If obj_listener.get_collection(areas_collection).Count > 0 Then
        obj_area_step.obj_area = obj_listener.get_collection(areas_collection)(rg_record.Offset(0, INT_OFFSET_AREA_ID).Value)
    End If

    obj_listener.get_collection(str_module).add obj_area_step, obj_area_step.str_id
Next obj_listener

End Function

Private Sub Class_Initialize()

str_module = new_const_mdl_id.STR_MDL_AREA_STEP
areas_collection = new_const_mdl_id.STR_MDL_AREA

Set col_listeners = New Collection

Set single_data_provider = New FileExcelDataProvider

INT_OFFSET_ID = 0
INT_OFFSET_AREA_ID = 1
INT_OFFSET_PLACE_FROM = 2
INT_OFFSET_PLACE_TO = 3

End Sub
