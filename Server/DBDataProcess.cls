VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBDataProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  ' pallet level
Public str_pallet As String
Public str_material As String
Public str_material_vendor As String
Public str_material_bin_storage_group As String
Public str_source As String
Public str_source_type As String
Public str_destination As String
Public str_destination_type As String
Public lng_stock_unit As Long
  ' process level
Public str_creation_id As String
Public str_version As String
  ' step level
Public str_user As String
Public str_date_start As String
Public str_date_end As String
Public str_bin_from As String
Public str_bin_to As String
Public str_transaction_type_start As String
Public str_transaction_type_end As String
'Public str_order_status As String
Public byte_process_status As Byte
Public byte_process_step_status As Byte

Public Property Get bool_is_single_transaction() As Boolean
    bool_is_single_transaction = str_date_end = ""
End Property

