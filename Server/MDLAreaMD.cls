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

Public str_module As String

Public INT_ID_OFFSET As Integer
Public INT_NAME_OFFSET As Integer

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
Dim obj_area As Area
Set obj_area = New Area

obj_area.str_id = rg_record.Offset(0, Me.INT_ID_OFFSET).Value
obj_area.str_name = rg_record.Offset(0, Me.INT_NAME_OFFSET).Value

col_areas.add obj_area

For Each obj_listener In Me.col_listeners
    obj_listener.process_record str_module, obj_area
Next obj_listener

End Function

Public Property Get get_mdl_id() As String

get_mdl_id = str_module

End Property

Private Sub Class_Initialize()

str_module = new_const_mdl_id.STR_MDL_AREA

Set col_listeners = New Collection
Set col_areas = New Collection

Set single_data_provider = New FileExcelDataProvider

INT_ID_OFFSET = 0
INT_NAME_OFFSET = 1

End Sub
