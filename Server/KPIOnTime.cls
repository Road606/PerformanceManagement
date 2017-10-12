VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIOnTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String

Public str_abbrev As String
Public str_type As String

Public dbl_target As Double
Public str_name As String

Public private_obj_process As ProcessMaster
Public col_process_versions As Collection
'Public col_limits As Collection
Public col_limits_process As Collection

Public Property Get obj_process() As ProcessMaster
    Set obj_process = private_obj_process
End Property

Public Property Let obj_process(obj_process As ProcessMaster)
    Set private_obj_process = obj_process
    obj_process.add_kpi Me
End Property

Public Function add_process_version(obj_process_version As ProcessMasterVersion)
    col_process_versions.add obj_process_version, obj_process_version.str_id
    obj_process_version.add_kpi Me
End Function

'Public Function add_limit(obj_limit As KPIOnTimeLimit)
'    col_limits.add obj_limit, obj_limit.str_id
'End Function

Public Function add_limit_process(obj_limit_process As KPIOnTimeLimitProc)
    col_limits_process.add obj_limit_process, obj_limit_process.str_id
End Function

Private Sub Class_Initialize()
    Set col_process_versions = New Collection
    'Set col_limits = New Collection
    Set col_limits_process = New Collection
End Sub
