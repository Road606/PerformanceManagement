VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String

Public str_material As String
Public str_material_vendor As String
Public str_material_bin_storage_group As String
Public str_source As String
Public str_destination As String

Public str_source_destination As String

Public lng_stock_unit As Long

'Public str_date As String
'Public str_bin As String

'Public str_process_master_id As String
'Public int_process_step_order As Integer
'Public byte_process_status As Byte

Public private_obj_process As process

Public Property Get obj_process() As process
    Set obj_process = private_obj_process
End Property

Public Property Let obj_process(obj_process As process)
    Set private_obj_process = obj_process
    Set obj_process.obj_pallet = Me
End Property

