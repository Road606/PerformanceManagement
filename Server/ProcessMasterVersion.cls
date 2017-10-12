VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessMasterVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public STEP_COL_ID As String

Public str_id As String
Public str_name As String
Private private_obj_master As ProcessMaster

Public col_steps As Collection
Public col_kpis As Collection
Public col_vendors As Collection

Public Property Get obj_master() As ProcessMaster
    Set obj_master = private_obj_master
End Property

Public Property Let obj_master(obj_master As ProcessMaster)
    Set private_obj_master = obj_master
    obj_master.add_version Me
End Property

Public Function add_step(obj_step As ProcessMasterStep)
    col_steps.add obj_step, retrieve_step_id(obj_step.str_place_from, obj_step.str_place_to)
End Function

Public Function get_step(str_place_from As String, str_place_to As String) As ProcessMasterStep
    Set get_step = col_steps.Item(retrieve_step_id(str_place_from, str_place_to))
End Function

Private Function retrieve_step_id(str_place_from As String, str_place_to As String) As String
    retrieve_step_id = str_place_from & STEP_COL_ID & str_place_to
    'retrieve_step_id = str_place_from & STEP_COL_ID & str_place_to & STEP_COL_ID & str_version
'    rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_FROM_TYPE), _
'        rg_record.Offset(0, db_history_record.STR_HISTORY_FLOW_OFFSET_BIN_TO_TYPE))
End Function

Public Function add_kpi(obj_kpi As Object)
    col_kpis.add obj_kpi, obj_kpi.str_id
End Function

Private Sub Class_Initialize()
    STEP_COL_ID = "-"

    Set col_steps = New Collection
    Set col_kpis = New Collection
    Set col_vendors = New Collection
End Sub

