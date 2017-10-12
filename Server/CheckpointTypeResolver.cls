VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckpointTypeResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function retrieve_building_for_checkpoint(str_source As String, str_source_type As String, _
        str_destination As String, str_destination_type As String, obj_checkpoint As KPIOnTimeLimitCheckpoint)
    Dim str_building_source As String
    Dim str_building_destination As String
    
    ' resolve source building
    str_building_source = retrieve_building(str_source, str_source_type)
    str_building_destination = retrieve_building(str_destination, str_destination_type)
    
    Select Case obj_checkpoint.str_type
        Case const_place_type.STR_PROCESS_SOURCE
            retrieve_building_for_checkpoint = str_building_source
        Case const_place_type.STR_PROCESS_TRANSITION
            ' # implement - not implemented yet, maybe never will be because data will be taken from table pallet path
        Case const_place_type.STR_PROCESS_DESTINATION
            retrieve_building_for_checkpoint = str_building_destination
    End Select
End Function

Public Function retrieve_building(str_data As String, str_data_type As String) As String
    Select Case str_data_type
        Case const_place_type.str_bin
            retrieve_building = bin.get_building(str_data)
        Case const_place_type.str_wc
            retrieve_building = wc.get_building(str_data)
        Case const_place_type.STR_IMPLICIT
            retrieve_building = str_data
    End Select
End Function
