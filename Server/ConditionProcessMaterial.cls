VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConditionProcessMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_condition_id As String

Public str_vendor As String
Public str_process_id As String
Public str_process_version_id As String


Public Function is_match(str_input_vendor As String, str_input_process_id As String, str_input_process_version_id As String) As Boolean
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
    ' check process
    If is_match Then
        obj_regex.Pattern = str_process_id
        is_match = obj_regex.test(str_input_process_id)
    End If
        
    ' check process_version
    If is_match Then
        obj_regex.Pattern = str_process_version_id
        is_match = obj_regex.test(str_input_process_version_id)
    End If
End Function
