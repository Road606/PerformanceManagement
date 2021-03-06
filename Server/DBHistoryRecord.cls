VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBHistoryRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_transaction_type_started As String
Public str_transaction_type_finished As String
Public str_transaction_started As String
Public str_transaction_finished As String
Public str_transaction_code As String
Public str_user_name As String
Public str_combi_vhu_from As String
Public str_combi_vhu_to As String
Public str_app_vhu As String
Public lng_qty_from As Long
Public lng_qty_to As Long
Public dbl_pallet_utilization As Double
Public str_material As String
Public str_bin_from As String
Public str_bin_to As String
Public str_hu_status_from As String
Public str_hu_status_to As String
Public str_hu_source As String
Public str_hu_destination As String
Public str_stock_type_from As String
Public str_stock_type_to As String
Public str_storage_group_material As String

Public str_gr_datetime As String

Public str_gr_type As String
Public str_gr_vendor As String
Public str_gr_order As String
Public str_gr_wc As String

Public str_order_delivery As String
Public str_wc_ship_to As String
Public str_machine_transport_ref As String

Public str_task_list_type As String

Public str_shipping_type As String

Public Function is_single_transaction() As Boolean
    is_single_transaction = str_transaction_finished = ""
End Function

Public Function is_start_transaction() As Boolean
    is_start_transaction = (str_combi_vhu_from = "") And (str_combi_vhu_to <> "")
End Function

Public Function is_final_transaction() As Boolean
    is_final_transaction = (str_combi_vhu_to = "") And (str_combi_vhu_from <> "")
End Function
