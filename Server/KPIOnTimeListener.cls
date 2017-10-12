VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIOnTimeListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STR_CLASS_NAME As String = "KPIOnTimeListener" ' # implement should be changed to KPIOnTimeResult

Public str_id As String
Public obj_kpi As Object

Public col_data_process_listeners As Collection
Public col_kpi_on_time_pallet_listeners As Collection

Public obj_md_dashboard As MDDashboard1
'Public obj_mdl_kpi_on_time_pallet As MDLKPIOnTimePallet
'Public obj_mdl_kpi_on_time_result As MDLKPIOnTimeResult

' data process listening
  ' # method is called from MDLProcessDashboardListener
Public Function create_intervals(obj_period_start As Date, obj_period_end As Date)
    Dim obj_listener As Object
    
    For Each obj_listener In col_data_process_listeners
        obj_listener.create_intervals obj_period_start, obj_period_end
    Next
End Function

  ' # method is called from MDLProcessDashboardListener
Public Function process_db_data_process_record(obj_data_process As DBDataProcess)
    Dim obj_listener As Object
    
    For Each obj_listener In col_data_process_listeners
        obj_listener.process_db_data_process_record obj_data_process
    Next
End Function

  ' # method is called from MDLProcessDashboardListener
Public Function save(obj_provider_info As FileExcelDataProviderInfo)
    Dim obj_listener As Object
    
    For Each obj_listener In col_data_process_listeners
        obj_listener.save obj_provider_info
    Next
End Function

' KPI pallet listening
Public Function kpi_pallet_loading_data_is_about_to_start(obj_date_start As Date, obj_date_end As Date)
    Dim obj_listener As Object
    
    For Each obj_listener In col_kpi_on_time_pallet_listeners
        obj_listener.kpi_pallet_loading_data_is_about_to_start obj_date_start, obj_date_end
    Next
End Function

' # method is called from MDLProcessDashboardListener - KPI pallet
Public Function kpi_pallet_process_record(obj_kpi_on_time_pallet As KPIOnTimePallet)
    Dim obj_listener As Object
    
    For Each obj_listener In col_kpi_on_time_pallet_listeners
        obj_listener.kpi_pallet_process_record obj_kpi_on_time_pallet
    Next
End Function


Private Sub Class_Initialize()
    Set col_data_process_listeners = New Collection
    Set col_kpi_on_time_pallet_listeners = New Collection
End Sub
