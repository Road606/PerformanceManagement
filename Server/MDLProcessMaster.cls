VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLProcessMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_module As String

Public single_data_provider As FileExcelDataProvider

Public col_listeners As Collection


Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function

Public Function load_data()
    single_data_provider.add_listener Me
    single_data_provider.load_data
End Function

Public Function load_record(rg_record As Range)
    Dim obj_process_master As ProcessMaster
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
    
        Set obj_process_master = New ProcessMaster
        obj_process_master.str_process_id = rg_record.Offset(0, new_db_process_master.INT_OFFSET_CREATION_ID).Value
        obj_process_master.str_type = rg_record.Offset(0, new_db_process_master.INT_OFFSET_TYPE).Value
        obj_process_master.str_subtype = rg_record.Offset(0, new_db_process_master.INT_OFFSET_SUBTYPE).Value
        obj_process_master.str_version_determinant = rg_record.Offset(0, new_db_process_master.INT_OFFSET_VERSION_DETERMINANT).Value
        
        obj_listener.get_collection(str_module).add obj_process_master, obj_process_master.str_process_id
    Next obj_listener
End Function

Private Sub Class_Initialize()

str_module = new_const_mdl_id.STR_MDL_PROCESS_MASTER

Set col_listeners = New Collection

Set single_data_provider = New FileExcelDataProvider

End Sub
