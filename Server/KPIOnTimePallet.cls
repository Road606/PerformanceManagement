VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIOnTimePallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String

Public str_process_id As String
Public str_process_version_id As String
Public str_process_version_name As String

Public str_material_vendor As String
Public str_material_bin_storage_group As String

Public str_building_source As String
Public str_building_destination As String

Public str_place_building As String

Public str_place_from As String
Public str_place_to As String

Public str_date_start As String
Public str_date_end As String
Public str_date_last_checkpoint As String

Public obj_checkpoint As KPIOnTimeLimitCheckpoint
Public obj_kpi_interval As KPIOnTimeInterval

' # implement as byte
Public str_process_status As String
Public byte_process_status As Byte
Public private_bool_status_ok As Boolean

Public Property Get bool_status_ok() As Boolean
    bool_status_ok = retrieve_actual_duration <= obj_checkpoint.obj_limit
End Property

Public Function retrieve_actual_duration() As Date
    retrieve_actual_duration = CDate(str_date_end) - CDate(str_date_start)
End Function

Public Function create_copy() As KPIOnTimePallet
    Set create_copy = New KPIOnTimePallet
    create_copy.str_id = str_id
    
    create_copy.str_process_id = str_process_id
    create_copy.str_process_version_id = str_process_version_id
    create_copy.str_process_version_name = str_process_version_name
    
    create_copy.str_place_from = str_place_from
    create_copy.str_place_to = str_place_to
    create_copy.str_place_building = str_place_building
    
    create_copy.str_date_start = str_date_start
    create_copy.str_date_end = str_date_end
    create_copy.str_date_last_checkpoint = str_date_last_checkpoint
    
    create_copy.str_material_vendor = str_material_vendor
    create_copy.str_material_bin_storage_group = str_material_bin_storage_group
    
    create_copy.str_building_source = str_building_source
    create_copy.str_building_destination = str_building_destination
    
    Set create_copy.obj_checkpoint = obj_checkpoint
    
    create_copy.byte_process_status = byte_process_status
End Function
