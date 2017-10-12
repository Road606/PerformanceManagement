VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZKPIOnTimePalletStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String

Public str_process_version_id As String

Public str_mat_grp_wh As String
Public str_source As String

Public str_place_building As String

Public str_place_from As String
Public str_place_to As String

Public str_date_start As String
Public str_date_end As String

Public bool_status_ok As Boolean

Public Function retrieve_actual_duration() As Date
    retrieve_actual_duration = CDate(str_date_end) - CDate(str_date_start)
End Function
