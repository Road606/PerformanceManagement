VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLKPIMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_module As String
Public str_process_collection As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_PROCESS_ID As Integer
Public INT_OFFSET_PROCESS_VERSION As Integer
Public INT_OFFSET_ABBREV As Integer
Public INT_OFFSET_TYPE As Integer
Public INT_OFFSET_KPI_TARGET As Integer
Public INT_OFFSET_CS_NAME As Integer

Public STR_TYPE_ON_TIME As String
Public STR_TYPE_FLOW As String

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
Dim obj_kpi As Object

For Each obj_listener In col_listeners
    On Error GoTo INFO_NEW_KPI_SETTING
        Set obj_kpi = obj_listener.get_collection(str_module)(rg_record.Offset(0, INT_OFFSET_ID).Value)
    On Error GoTo 0

    If Not obj_kpi.obj_process Is Nothing Then
        obj_kpi.add_process_version obj_kpi.obj_process.col_versions(rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION).Value)
    End If

    Exit Function

INFO_NEW_KPI_SETTING:
    Set obj_kpi = create_kpi(rg_record)
    obj_kpi.str_id = rg_record.Offset(0, INT_OFFSET_ID).Value
    obj_kpi.str_abbrev = rg_record.Offset(0, INT_OFFSET_ABBREV).Value
    obj_kpi.str_type = rg_record.Offset(0, INT_OFFSET_TYPE).Value
    obj_kpi.dbl_target = rg_record.Offset(0, INT_OFFSET_KPI_TARGET).Value
    obj_kpi.str_name = rg_record.Offset(0, INT_OFFSET_CS_NAME).Value

    If obj_listener.get_collection(str_process_collection).Count > 0 Then
        obj_kpi.obj_process = obj_listener.get_collection(str_process_collection)(rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value)
    End If


    obj_listener.get_collection(str_module).add obj_kpi, obj_kpi.str_id
Next obj_listener

End Function

Private Function create_kpi(rg_record) As Object
    Select Case rg_record.Offset(0, INT_OFFSET_TYPE).Value
        Case STR_TYPE_ON_TIME
            Set create_kpi = New KPIOnTime
        Case STR_TYPE_FLOW
    End Select
End Function

Private Sub Class_Initialize()

str_module = new_const_mdl_id.STR_MDL_KPI
str_process_collection = new_const_mdl_id.STR_MDL_PROCESS_MASTER

Set col_listeners = New Collection

Set single_data_provider = New FileExcelDataProvider

    INT_OFFSET_ID = 0
    INT_OFFSET_PROCESS_ID = 1
    INT_OFFSET_PROCESS_VERSION = 2
    INT_OFFSET_ABBREV = 3
    INT_OFFSET_TYPE = 4
    INT_OFFSET_KPI_TARGET = 5
    INT_OFFSET_CS_NAME = 6
    
    STR_TYPE_ON_TIME = "ON_TIME"
    STR_TYPE_FLOW = "FLOW"

End Sub
