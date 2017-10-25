VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLProcessVersionMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_module As String
Public str_process_collection As String

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
    Dim obj_process_version As ProcessMasterVersion
    Dim obj_listener As Object
    
    For Each obj_listener In col_listeners
    
        Set obj_process_version = New ProcessMasterVersion
        obj_process_version.str_id = rg_record.Offset(0, new_db_process_master_version.INT_OFFSET_VERSION).Value
        obj_process_version.str_name = rg_record.Offset(0, new_db_process_master_version.INT_OFFSET_NAME).Value
    
        If obj_listener.get_collection(str_process_collection).Count > 0 Then
        ' # implement - INT_OFFSET_PROCESS_ID should be loaded from config not from module
            obj_process_version.obj_master = obj_listener.get_collection(str_process_collection).Item( _
                rg_record.Offset(0, new_db_process_master_version.INT_OFFSET_PROCESS_ID).Value)
        End If
            obj_listener.get_collection(str_module).add obj_process_version
        
    Next obj_listener
    
End Function

Private Sub Class_Initialize()

str_module = new_const_mdl_id.STR_MDL_PROCESS_VERSION_MASTER
str_process_collection = new_const_mdl_id.STR_MDL_PROCESS_MASTER

Set col_listeners = New Collection

Set single_data_provider = New FileExcelDataProvider

End Sub

