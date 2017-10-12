VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListenerHistoryToProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function before_process_history_record(obj_data_provider_info As FileExcelDataProviderInfo)
    Set new_ctrl_pallet.obj_data_provider_info = obj_data_provider_info
End Function

Public Function process_history_record(obj_record As DBHistoryRecord)
    new_ctrl_pallet.process_record obj_record
End Function

'Public Function process_record(obj_record As DBDataProcess)
'    new_ctrl_pallet.process_record obj_record
'End Function
