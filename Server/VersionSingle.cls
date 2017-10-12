VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionSingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_version_id As String
Public str_source As String
Public str_source_type As String
Public str_destination As String
Public str_destination_type As String

Public Function init(obj_record As DBHistoryRecord)
    str_version_id = new_ctrl_process_master_version.STR_VERSION_DEFAULT
End Function


