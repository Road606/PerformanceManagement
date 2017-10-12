VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIOnTimeLimitMatGrpWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String
Public col_limits_source As Collection

Public Function add_limit(obj_limit As KPIOnTimeLimitSource)
    col_limits_source.add obj_limit, obj_limit.str_id
End Function

Private Sub Class_Initialize()
    Set col_limits_source = New Collection
End Sub

