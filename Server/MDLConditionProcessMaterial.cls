VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLConditionProcessMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_module As String

Public INT_OFFSET_CONDITION_ID As Integer
Public INT_OFFSET_VENDOR As Integer
Public INT_OFFSET_PROCESS_ID As Integer
Public INT_OFFSET_PROCESS_VERSION_ID As Integer

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
    Dim obj_vendor As Vendor
    Dim obj_condition As ConditionProcessMaterial
     Dim obj_listener As Object

    For Each obj_listener In col_listeners
        On Error GoTo INFO_NEW_VENDOR
        Set obj_vendor = obj_listener.get_collection(str_module)(CStr(rg_record.Offset(0, INT_OFFSET_VENDOR).Value))
        On Error GoTo 0
    
        Set obj_condition = New ConditionProcessMaterial
        obj_condition.str_condition_id = rg_record.Offset(0, INT_OFFSET_CONDITION_ID).Value
        obj_condition.str_vendor = obj_vendor.str_id
        obj_condition.str_process_id = rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value
        obj_condition.str_process_version_id = rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION_ID).Value
    ' add to vendor
        obj_vendor.add_condition_process_material obj_condition
    Next obj_listener
    
    Exit Function
INFO_NEW_VENDOR:
    Set obj_vendor = create_vendor(rg_record, obj_listener)
    Resume Next
End Function

Public Function create_vendor(rg_record As Range, obj_listener As Object) As Vendor
    Set create_vendor = New Vendor
    create_vendor.str_id = rg_record.Offset(0, INT_OFFSET_VENDOR).Value
    
    obj_listener.get_collection(str_module).add create_vendor, create_vendor.str_id
End Function

Private Sub Class_Initialize()

str_module = new_const_mdl_id.STR_MDL_CONDITION_PROCESS_MATERIAL

Set col_listeners = New Collection

Set single_data_provider = New FileExcelDataProvider

    INT_OFFSET_CONDITION_ID = 0
    INT_OFFSET_VENDOR = 1
    INT_OFFSET_PROCESS_ID = 2
    INT_OFFSET_PROCESS_VERSION_ID = 3

End Sub


