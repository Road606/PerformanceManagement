VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String
Public int_lower_limit As Integer
Public int_upper_limit As Integer
Public int_minute As Integer
Public int_hour_offset As Integer
Public int_is_complete As Integer

Public private_obj_process As ProcessMaster

Public Property Get obj_process() As ProcessMaster
    Set obj_process = private_obj_process
End Property

Public Property Let obj_process(obj_process As ProcessMaster)
    Set private_obj_process = obj_process
    obj_process.add_interval Me
End Property
