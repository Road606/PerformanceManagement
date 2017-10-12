VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public obj_process As process
Public obj_master_step As ProcessMasterStep

Public str_user As String

Public str_date_start As String
Public str_date_end As String

Public str_bin_from As String
Public str_bin_to As String

Public str_place_from As String
Public str_place_to As String

Public str_transaction_type_start As String
Public str_transaction_type_end As String

Public str_order_status As String

Public byte_status As Byte
Public byte_process_status As Byte

Public Function get_master() As ProcessMasterStep
    If obj_master_step Is Nothing Then
        On Error GoTo INFO_MASTER_STEP_NOT_FOUND
        Set obj_master_step = obj_process.obj_master_version.get_step( _
            str_place_from, str_place_to)
        On Error GoTo 0
    End If
    
    Set get_master = obj_master_step
    Exit Function
INFO_MASTER_STEP_NOT_FOUND:
    Set obj_master_step = obj_process.obj_master_version.get_step("", "") ' get default
    Resume Next
End Function

Public Function reevaluate_step_status()
    If str_date_end = "" Then
        byte_status = new_db_process_step.BYTE_OPEN
    Else
        byte_status = new_db_process_step.BYTE_CLOSED
    End If
End Function

Public Function to_string_status() As String
    If byte_status = new_db_process_step.BYTE_CLOSED Then
        to_string_status = new_db_process_step.STR_CLOSED
    Else
        to_string_status = new_db_process_step.STR_OPEN
    End If
End Function
