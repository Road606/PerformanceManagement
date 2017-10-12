VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIOnTimeLimitProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String
Public col_limits_version As Collection

Public private_obj_kpi As KPIOnTime

Public Property Get obj_kpi() As KPIOnTime
    Set obj_kpi = private_obj_kpi
End Property

Public Property Let obj_kpi(obj_kpi As KPIOnTime)
    Set private_obj_kpi = obj_kpi
    obj_kpi.add_limit_process Me
End Property

Public Function add_limit(obj_limit As KPIOnTimeLimitProcVersion)
    col_limits_version.add obj_limit, obj_limit.str_id
End Function

Private Sub Class_Initialize()
    Set col_limits_version = New Collection
End Sub

