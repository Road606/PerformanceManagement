VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String
Public str_name As String

Public col_material_group_wh_resolvers As Collection
Public col_conditions_process_material As Collection

Public Function add_material_group_wh_resolver(obj_material_group_wh_resolver As MaterialGroupWHResolver)
    col_material_group_wh_resolvers.add obj_material_group_wh_resolver
    obj_material_group_wh_resolver.str_vendor = str_id
End Function

Public Function retrieve_material_group_wh_resolver(str_bin_storage_group As String) As MaterialGroupWHResolver
    Dim obj_material_group_wh_resolver As MaterialGroupWHResolver
    
    For Each obj_material_group_wh_resolver In col_material_group_wh_resolvers
        If obj_material_group_wh_resolver.is_match(str_id, str_bin_storage_group) Then
            Set retrieve_material_group_wh_resolver = obj_material_group_wh_resolver
            Exit For
        End If
    Next
End Function

Public Function add_condition_process_material(obj_condition As ConditionProcessMaterial)
    col_conditions_process_material.add obj_condition
    obj_condition.str_vendor = str_id
End Function

Public Function retrieve_condition_process_material(str_process_id As String, str_process_version_id As String) As ConditionProcessMaterial
    Dim obj_condition As ConditionProcessMaterial
    
    For Each obj_condition In col_conditions_process_material
        If obj_condition.is_match(str_id, str_process_id, str_process_version_id) Then
            Set retrieve_condition_process_material = obj_condition
            Exit For
        End If
    Next
End Function

Private Sub Class_Initialize()
    Set col_material_group_wh_resolvers = New Collection
    Set col_conditions_process_material = New Collection
End Sub
