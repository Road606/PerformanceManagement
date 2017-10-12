VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaterialGroupWHResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_material_group_wh_id As String
Public str_vendor As String
Public str_bin_storage_group_pattern As String
Public str_cs_name As String

Public Function is_match(str_input_vendor As String, str_input_bin_storage_group_pattern As String) As Boolean
    Dim obj_regex As RegExp
    
    Set obj_regex = New RegExp
    With obj_regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With
    
    ' check vendor
    obj_regex.Pattern = str_vendor
    is_match = obj_regex.test(str_input_vendor)
        
    ' check bin storage group
    If is_match Then
        obj_regex.Pattern = str_bin_storage_group_pattern
        is_match = obj_regex.test(str_input_bin_storage_group_pattern)
    End If
End Function
