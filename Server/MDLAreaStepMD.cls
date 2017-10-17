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

Public INT_ID_OFFSET As Integer
Public INT_AREA_ID_OFFSET As Integer
Public INT_PLACE_FROM_OFFSET As Integer
Public INT_PLACE_TO_OFFSET As Integer

Public single_data_provider As FileExcelDataProvider

Public col_listeners As Collection
Public col_areas As Collection

Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function

Public Function load_data(col_area As Collection)
    Set col_areas = col_area
    single_data_provider.add_listener Me
    single_data_provider.load_data
End Function

Public Function load_record(rg_record As Range)
Dim obj_listener As Object
Dim obj_area_step As AreaStep

Set obj_area_step = New AreaStep
obj_area_step.str_id = rg_record.Offset(0, INT_ID_OFFSET).Value
obj_area_step.str_place_from = rg_record.Offset(0, INT_PLACE_FROM_OFFSET).Value
obj_area_step.str_place_to = rg_record.Offset(0, INT_PLACE_TO_OFFSET).Value

If col_areas.Count > 0 Then
        obj_area_step.obj_area = col_areas(rg_record.Offset(0, INT_AREA_ID_OFFSET).Value)
End If

For Each obj_listener In Me.col_listeners
    obj_listener.process_record str_module, obj_area_step
Next obj_listener

End Function

Private Sub Class_Initialize()

str_module = new_const_mdl_id.STR_MDL_AREA_STEP

Set col_listeners = New Collection

Set single_data_provider = New FileExcelDataProvider

INT_ID_OFFSET = 0
INT_AREA_ID_OFFSET = 1
INT_PLACE_FROM_OFFSET = 2
INT_PLACE_TO_OFFSET = 3

End Sub
