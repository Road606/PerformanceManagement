VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public obj_pallet As Pallet
Public obj_master_version As ProcessMasterVersion

Public str_source As String
Public str_source_type As String
Public str_destination As String
Public str_destination_type As String


Public obj_actual_step As ProcessStep
Public obj_steps As Collection

Public int_last_step_id As Integer

Public byte_status As Byte

Private Sub Class_Initialize()
    int_last_step_id = 0
    
    Set obj_steps = New Collection
End Sub

Public Function add_step(obj_step As ProcessStep)
'    ' evaluate if the step is in order
'    If obj_step.str_order_status = "" Then ' order status is not empty if master step was not found
'        If int_actual_step_id + 1 = obj_step.obj_master_step.int_order Then
'            obj_step.str_order_status = new_ctrl_process.STR_ORDER_STATUS_OK
'        ElseIf int_actual_step_id + 1 > obj_step.obj_master_step.int_order Then
'            obj_step.str_order_status = new_ctrl_process.STR_ORDER_STATUS_AHEAD
'        Else
'            obj_step.str_order_status = new_ctrl_process.STR_ORDER_STATUS_BACK
'        End If
'    End If
'
'    int_actual_step_id = obj_step.obj_master_step.int_order
    obj_steps.add obj_step ', obj_step.obj_master_step.str_place_to
    Set obj_step.obj_process = Me
End Function

'Public Function get_step(str_place_from As String) As ProcessStep
'    Set get_step = obj_steps.Item(str_place_from)
'End Function

Public Function has_open_step() As Boolean
    has_open_step = True

    If obj_actual_step Is Nothing Then
        has_open_step = False
    End If
End Function
