VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLPlaceMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_module As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_NAME As Integer

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
Dim obj_place As Place
Set obj_place = New Place

obj_place.str_id = rg_record.Offset(0, Me.INT_OFFSET_ID).Value
obj_place.str_name = rg_record.Offset(0, Me.INT_OFFSET_NAME).Value

For Each obj_listener In Me.col_listeners
    obj_listener.get_collection(str_module).add obj_place, obj_place.str_id
Next obj_listener


End Function

Private Sub Class_Initialize()

str_module = new_const_mdl_id.STR_MDL_PLACE

Set col_listeners = New Collection

Set single_data_provider = New FileExcelDataProvider

INT_OFFSET_ID = 0
INT_OFFSET_NAME = 1

End Sub
