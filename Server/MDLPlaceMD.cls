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

Public INT_ID_OFFSET As Integer
Public INT_NAME_OFFSET As Integer

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

obj_place.str_id = rg_record.Offset(0, Me.INT_ID_OFFSET).Value
obj_place.str_name = rg_record.Offset(0, Me.INT_NAME_OFFSET).Value

For Each obj_listener In Me.col_listeners
    obj_listener.process_record str_module, obj_place
Next obj_listener


End Function

Private Sub Class_Initialize()

str_module = new_const_mdl_id.STR_MDL_PLACE

Set col_listeners = New Collection

Set single_data_provider = New FileExcelDataProvider

INT_ID_OFFSET = 0
INT_NAME_OFFSET = 1

End Sub
