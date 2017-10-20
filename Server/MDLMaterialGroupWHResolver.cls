VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLMaterialGroupWHResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_module As String

Public INT_OFFSET_MATERIAL_GROUP_WH As Integer
Public INT_OFFSET_VENDOR As Integer
Public INT_OFFSET_BIN_STORAGE_GROUP_PATTERN As Integer
Public INT_OFFSET_CS_NAME As Integer

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
    Dim obj_vendor As Vendor
    Dim obj_material_group_wh_resolver As MaterialGroupWHResolver

    For Each obj_listener In col_listeners
        On Error GoTo INFO_NEW_VENDOR
            Set obj_vendor = obj_listener.get_collection(str_module)(rg_record.Offset(0, INT_OFFSET_VENDOR).Value)
        On Error GoTo 0
    
        Set obj_material_group_wh_resolver = New MaterialGroupWHResolver
        obj_material_group_wh_resolver.str_material_group_wh_id = rg_record.Offset(0, INT_OFFSET_MATERIAL_GROUP_WH).Value
        obj_material_group_wh_resolver.str_bin_storage_group_pattern = rg_record.Offset(0, INT_OFFSET_BIN_STORAGE_GROUP_PATTERN).Value
        ' add to vendor
        obj_vendor.add_material_group_wh_resolver obj_material_group_wh_resolver
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

str_module = new_const_mdl_id.STR_MDL_MATERIAL_GROUP_WH_RESOLVER

Set col_listeners = New Collection

Set single_data_provider = New FileExcelDataProvider

INT_OFFSET_MATERIAL_GROUP_WH = 0
INT_OFFSET_VENDOR = 1
INT_OFFSET_BIN_STORAGE_GROUP_PATTERN = 2
INT_OFFSET_CS_NAME = 3

End Sub
