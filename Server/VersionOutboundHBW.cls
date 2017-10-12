VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionOutboundHBW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public STR_VERSION_DEFAULT As String

Public str_version_id As String
Public str_source As String
Public str_source_type As String
Public str_destination As String
Public str_destination_type As String

Public Function init(obj_record As DBHistoryRecord)
    str_version_id = STR_VERSION_DEFAULT
'    Dim str_building_from As String
'    Dim str_building_to As String
'
'    str_building_from = bin.get_building(obj_record.str_bin_from)
'    str_building_to = wc.get_building(obj_record.str_wc_shipto)
'
'    Select Case str_building_from & str_building_to
'        Case "AA", "BB", "CC"
'            retrieve = new_ctrl_process_master_version.STR_VERSION_SINGLE
'        Case "BC", "CB"
'            retrieve = new_ctrl_process_master_version.STR_VERSION_SINGLE
'        Case "AB", "BA", "AC", "CA"
'            retrieve = new_ctrl_process_master_version.STR_VERSION_CROSS
'    End Select
End Function

Private Sub Class_Initialize()
    STR_VERSION_DEFAULT = "0"
End Sub
