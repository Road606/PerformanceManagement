VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessMasterStepOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public obj_version As ProcessMasterVersion

Public str_place_from As String
Public str_place_to As String

Public int_order As Integer ' create as collection

Public str_action_type As String
Public str_action_parameter_from As String
Public str_action_parameter_to As String

'Public bool_status As Boolean

'Public Property Get str_id() As String
'    str_id = int_order & str_place_from & str_place_to
'End Property
