VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public str_transaction_code As String
'Public STR_TRANSACTION_USER As String
'Public STR_TRANSACTION_TASK_LIST_TYPE As String

Public str_process_id As String
Public str_version_determinant As String

'Public col_creation_transactions As Collection
'Public col_creation_parameters As Collection

Public str_type As String
Public str_subtype As String

Public str_transaction_type_allowed As String

Public col_actions As Collection
Public col_versions As Collection
Public col_kpis As Collection
Public col_intervals As Collection

Public Function add_action(obj_action As ProcessMasterAction)
    col_actions.add obj_action, obj_action.str_id
End Function
'
Public Function get_action(str_action_id As String) As ProcessMasterAction
    Set get_action = col_actions.Item(str_action_id)
End Function

Public Function add_version(obj_version As ProcessMasterVersion)
    col_versions.add obj_version, obj_version.str_id
End Function

Public Function get_version(str_version_id As String) As ProcessMasterVersion
    Set get_version = col_versions.Item(str_version_id)
End Function

Public Function add_interval(obj_interval As Object)
    col_intervals.add obj_interval, obj_interval.str_id
End Function

Public Function add_kpi(obj_kpi As Object)
    col_kpis.add obj_kpi, obj_kpi.str_id
End Function

Private Sub Class_Initialize()
    Set col_versions = New Collection
    Set col_actions = New Collection
    Set col_kpis = New Collection
    Set col_intervals = New Collection
    'Set col_creation_parameters = New Collection
    
'    str_transaction_code = "transaction_code"
'    STR_TRANSACTION_USER = "user"
'    STR_TRANSACTION_TASK_LIST_TYPE = "task_list_type"
End Sub

