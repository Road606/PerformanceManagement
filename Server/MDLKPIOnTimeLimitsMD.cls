VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLKPIOnTimeLimitsMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_module As String

Public INT_OFFSET_ID As Integer
Public INT_OFFSET_KPI_ID As Integer
Public INT_OFFSET_PROCESS_ID As Integer
Public INT_OFFSET_PROCESS_VERSION_ID As Integer
Public INT_OFFSET_CHECKPOINT_PLACE As Integer
Public INT_OFFSET_CHECKPOINT_PLACE_TYPE As Integer
Public INT_OFFSET_CHECKPOINT_PLACE_NEXT As Integer
Public INT_OFFSET_MAT_GRP_WH_ID As Integer
Public INT_OFFSET_SOURCE As Integer
Public INT_OFFSET_LIMIT As Integer


Public single_data_provider As FileExcelDataProvider

Public col_listeners As Collection
'Public col_kpis As Collection
Public col_kpi_limits As Collection

Dim obj_kpi_last_limit As KPIOnTimeLimitCheckpoint


Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function

Public Function load_data()
    
    single_data_provider.add_listener Me
    single_data_provider.load_data
    
End Function

Public Function load_record(rg_record As Range)
Dim obj_listener As Object
Dim obj_limit_process As KPIOnTimeLimitProc
Dim obj_limit_process_version As KPIOnTimeLimitProcVersion
Dim obj_limit_mat_grp_wh As KPIOnTimeLimitMatGrpWH
Dim obj_limit_source As KPIOnTimeLimitSource
Dim obj_limit_checkpoint_place As KPIOnTimeLimitCheckpoint

For Each obj_listener In col_listeners
    On Error GoTo INFO_NEW_PROCESS
    Set obj_limit_process = col_kpi_limits(rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value)
    On Error GoTo 0
    
    On Error GoTo INFO_NEW_PROCESS_VERSION
    Set obj_limit_process_version = obj_limit_process.col_limits_version(rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION_ID).Value)
    On Error GoTo 0

    On Error GoTo INFO_NEW_MAT_GRP_WH
    Set obj_limit_mat_grp_wh = obj_limit_process_version.col_limits_mat_grp_wh(rg_record.Offset(0, INT_OFFSET_MAT_GRP_WH_ID).Value)
    On Error GoTo 0
    
    On Error GoTo INFO_NEW_SOURCE
    Set obj_limit_source = obj_limit_mat_grp_wh.col_limits_source(rg_record.Offset(0, INT_OFFSET_SOURCE).Value)
    On Error GoTo 0
    
    Set obj_limit_checkpoint_place = New KPIOnTimeLimitCheckpoint
    obj_limit_checkpoint_place.str_id = rg_record.Offset(0, INT_OFFSET_CHECKPOINT_PLACE).Value
    obj_limit_checkpoint_place.str_type = rg_record.Offset(0, INT_OFFSET_CHECKPOINT_PLACE_TYPE).Value
    obj_limit_checkpoint_place.str_next_checkpoint_id = rg_record.Offset(0, INT_OFFSET_CHECKPOINT_PLACE_NEXT).Value
    obj_limit_checkpoint_place.obj_limit = rg_record.Offset(0, INT_OFFSET_LIMIT).Value
    If Not obj_kpi_last_limit Is Nothing Then
        Set obj_limit_checkpoint_place.obj_previous_limit = obj_kpi_last_limit
        Set obj_kpi_last_limit.obj_next_limit = obj_limit_checkpoint_place
    End If
    
    Set obj_kpi_last_limit = obj_limit_checkpoint_place
    obj_limit_source.add_limit obj_limit_checkpoint_place
Next obj_listener
Exit Function
INFO_NEW_PROCESS:
    Set obj_limit_process = create_process(rg_record, obj_listener.get_collection(str_module))
    Set obj_kpi_last_limit = Nothing
    Resume Next
INFO_NEW_PROCESS_VERSION:
    Set obj_limit_process_version = create_process_version(rg_record, obj_limit_process)
    Set obj_kpi_last_limit = Nothing
    Resume Next
INFO_NEW_MAT_GRP_WH:
    Set obj_limit_mat_grp_wh = create_mat_grp_wh(rg_record, obj_limit_process_version)
    Set obj_kpi_last_limit = Nothing
    Resume Next
INFO_NEW_SOURCE:
    Set obj_limit_source = create_source(rg_record, obj_limit_mat_grp_wh)
    Set obj_kpi_last_limit = Nothing
    Resume Next


End Function

Public Function create_process(rg_record As Range, col_kpis As Collection) As KPIOnTimeLimitProc
    Set create_process = New KPIOnTimeLimitProc
    create_process.str_id = rg_record.Offset(0, INT_OFFSET_PROCESS_ID).Value
    
    If Not col_kpis Is Nothing Then
        create_process.obj_kpi = col_kpis(rg_record.Offset(0, INT_OFFSET_KPI_ID).Value)
    End If
    
    col_kpi_limits.add create_process, create_process.str_id
End Function

Public Function create_process_version(rg_record As Range, obj_limit_process As KPIOnTimeLimitProc) As KPIOnTimeLimitProcVersion
    Set create_process_version = New KPIOnTimeLimitProcVersion
    create_process_version.str_id = rg_record.Offset(0, INT_OFFSET_PROCESS_VERSION_ID).Value
    
    obj_limit_process.add_limit create_process_version
End Function

Public Function create_mat_grp_wh(rg_record As Range, obj_limit_process_version As KPIOnTimeLimitProcVersion) As KPIOnTimeLimitMatGrpWH
    Set create_mat_grp_wh = New KPIOnTimeLimitMatGrpWH
    create_mat_grp_wh.str_id = rg_record.Offset(0, INT_OFFSET_MAT_GRP_WH_ID).Value
    
    obj_limit_process_version.add_limit create_mat_grp_wh
End Function

Public Function create_source(rg_record As Range, obj_limit_mat_grp_wh As KPIOnTimeLimitMatGrpWH) As KPIOnTimeLimitSource
    Set create_source = New KPIOnTimeLimitSource
    create_source.str_id = rg_record.Offset(0, INT_OFFSET_SOURCE).Value
    
    obj_limit_mat_grp_wh.add_limit create_source
End Function


Private Sub Class_Initialize()

str_module = new_const_mdl_id.STR_MDL_KPI_ONTIME_LIMITS

Set col_listeners = New Collection
Set col_kpi_limits = New Collection

Set single_data_provider = New FileExcelDataProvider

    INT_OFFSET_ID = 0
    INT_OFFSET_KPI_ID = 1
    INT_OFFSET_PROCESS_ID = 2
    INT_OFFSET_PROCESS_VERSION_ID = 3
    INT_OFFSET_CHECKPOINT_PLACE = 4
    INT_OFFSET_CHECKPOINT_PLACE_TYPE = 5
    INT_OFFSET_CHECKPOINT_PLACE_NEXT = 6
    INT_OFFSET_MAT_GRP_WH_ID = 7
    INT_OFFSET_SOURCE = 8
    INT_OFFSET_LIMIT = 9
    
End Sub
