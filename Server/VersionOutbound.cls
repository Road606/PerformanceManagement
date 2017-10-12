VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionOutbound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public STR_VERSION_SINGLE As String
Public STR_VERSION_CROSS As String

Public str_version_id As String
Public str_source As String
Public str_source_type As String
Public str_destination As String
Public str_destination_type As String


Public Function init(obj_record As DBHistoryRecord)
    Dim str_building_from As String
    Dim str_building_to As String
    Dim str_module As String
    Dim str_sub_version As String
    
    ' from
    str_source = obj_record.str_bin_from
    str_source_type = const_place_type.str_bin
    str_building_from = bin.get_building(obj_record.str_bin_from)
    
    'to
    str_destination = obj_record.str_machine_transport_ref
    str_destination_type = const_place_type.str_bin
    str_building_to = bin.get_building(obj_record.str_machine_transport_ref)
    str_module = wc.get_module(obj_record.str_wc_ship_to)
    
    Select Case str_building_from & str_building_to
        Case "AA", "BB", "CC"
            str_sub_version = STR_VERSION_SINGLE 'new_ctrl_process_master_version.STR_VERSION_SINGLE
        Case "BC", "CB"
            str_sub_version = STR_VERSION_SINGLE  'new_ctrl_process_master_version.STR_VERSION_SINGLE
        Case "AB", "BA", "AC", "CA"
            str_sub_version = STR_VERSION_CROSS
    End Select
    
    str_version_id = str_module & "-" & str_sub_version
End Function

Private Sub Class_Initialize()
    STR_VERSION_SINGLE = "1"
    STR_VERSION_CROSS = "2"
End Sub
