VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDDashboard1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public STR_WS_NAME_MASTER_PROCESS As String
Public STR_WS_NAME_MASTER_PROCESS_VERSION As String
Public STR_WS_NAME_MASTER_PROCESS_INTERVAL As String

Public STR_WS_NAME_AREA_MD As String
Public STR_WS_NAME_AREA_STEP_MD As String

Public STR_WS_NAME_PLACE_MD As String

Public STR_WS_NAME_KPI As String
Public STR_WS_NAME_KPI_ON_TIME_LIMITS As String

Public STR_WS_NAME_MATERIAL_GROUP_WH_RESOLVER As String

Public STR_WS_NAME_CONDITION_PROCESS_MATERIAL As String

Public col_processes As Collection
Public col_process_versions As Collection
Public col_area_steps As Collection
Public col_places As Collection
Public col_kpis As Collection
Public col_material_group_wh_resolvers As Collection
Public col_conditions_process_material As Collection

Public Function load()
    load_process
    load_area_steps
    load_places
    load_kpis
    load_material_group_wh_resolvers
    load_conditions_process_material
End Function

Private Function load_process()
    Dim obj_mdl_process As MDLProcessMaster
    Dim obj_mdl_process_version As MDLProcessVersionMaster
    Dim obj_mdl_process_interval As MDLProcessIntervalMD
    Dim obj_data_provider As FileExcelDataProvider

    Set obj_mdl_process = New MDLProcessMaster
    obj_mdl_process.init
    obj_mdl_process.STR_WS_NAME = STR_WS_NAME_MASTER_PROCESS
    
    Set obj_mdl_process_version = New MDLProcessVersionMaster
    'obj_mdl_process_version.init
    obj_mdl_process_version.STR_WS_NAME = STR_WS_NAME_MASTER_PROCESS_VERSION
    
    Set obj_mdl_process_interval = New MDLProcessIntervalMD
    Set obj_data_provider = New FileExcelDataProvider
    obj_data_provider.str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_INTERNAL
    obj_data_provider.STR_WS_NAME = STR_WS_NAME_MASTER_PROCESS_INTERVAL
    obj_data_provider.add_listener obj_mdl_process_interval
    Set obj_mdl_process_interval.obj_data_provider = obj_data_provider
    
    obj_mdl_process.load_data
    obj_mdl_process_version.load_data obj_mdl_process.col_process_masters
    Set obj_mdl_process_interval.col_processes = obj_mdl_process.col_process_masters
    obj_mdl_process_interval.load
    
    Set col_processes = obj_mdl_process.col_process_masters
    Set col_process_versions = obj_mdl_process_version.col_versions
    
    Set obj_mdl_process = Nothing
    Set obj_mdl_process_version = Nothing
    Set obj_mdl_process_interval = Nothing
    Set obj_data_provider = Nothing
End Function

Private Function load_area_steps()
    Dim obj_mdl_area As MDLAreaMD
    Dim obj_mdl_area_step As MDLAreaStepMD

    Set obj_mdl_area = New MDLAreaMD
    obj_mdl_area.init
    obj_mdl_area.STR_WS_NAME = STR_WS_NAME_AREA_MD '"db.md.area"
    
    Set obj_mdl_area_step = New MDLAreaStepMD
    obj_mdl_area_step.init
    obj_mdl_area_step.STR_WS_NAME = STR_WS_NAME_AREA_STEP_MD '"db.md.area.step"
    
    obj_mdl_area.load_data
    obj_mdl_area_step.load_data obj_mdl_area.col_areas
    
    Set col_area_steps = obj_mdl_area_step.col_area_steps
    
    Set obj_mdl_area = Nothing
    Set obj_mdl_area_step = Nothing
End Function

Private Function load_places()
    Dim obj_mdl_place As MDLPlaceMD

    Set obj_mdl_place = New MDLPlaceMD
    obj_mdl_place.init
    obj_mdl_place.STR_WS_NAME = STR_WS_NAME_PLACE_MD
    obj_mdl_place.load_data
    
    Set col_places = obj_mdl_place.col_places
    
    Set obj_mdl_place = Nothing
End Function

Public Function load_kpis()
    Dim obj_mdl_kpi As MDLKPIMD
    Dim obj_mdl_kpi_limit As MDLKPIOnTimeLimitsMD

    Set obj_mdl_kpi = New MDLKPIMD
    obj_mdl_kpi.init
    obj_mdl_kpi.STR_WS_NAME = STR_WS_NAME_KPI
    obj_mdl_kpi.load_data col_processes
    
    Set obj_mdl_kpi_limit = New MDLKPIOnTimeLimitsMD
    'obj_mdl_kpi_limit.init
    obj_mdl_kpi_limit.STR_WS_NAME = STR_WS_NAME_KPI_ON_TIME_LIMITS
    obj_mdl_kpi_limit.load_data obj_mdl_kpi.col_kpis
    
    Set col_kpis = obj_mdl_kpi.col_kpis
End Function

Public Function load_material_group_wh_resolvers()
    Dim obj_mdl_material_group_wh_resolver As MDLMaterialGroupWHResolver
    
    Set obj_mdl_material_group_wh_resolver = New MDLMaterialGroupWHResolver
    obj_mdl_material_group_wh_resolver.STR_WS_NAME = STR_WS_NAME_MATERIAL_GROUP_WH_RESOLVER
    obj_mdl_material_group_wh_resolver.load_data
    Set col_material_group_wh_resolvers = obj_mdl_material_group_wh_resolver.col_vendors
End Function

Public Function load_conditions_process_material()
    Dim obj_mdl_condition_process_material As MDLConditionProcessMaterial
    
    Set obj_mdl_condition_process_material = New MDLConditionProcessMaterial
    obj_mdl_condition_process_material.STR_WS_NAME = STR_WS_NAME_CONDITION_PROCESS_MATERIAL
    obj_mdl_condition_process_material.load_data
    Set col_conditions_process_material = obj_mdl_condition_process_material.col_vendors
End Function

Private Sub Class_Initialize()
    STR_WS_NAME_MASTER_PROCESS = "data.master.process"
    STR_WS_NAME_MASTER_PROCESS_VERSION = "data.master.process.version"
    STR_WS_NAME_MASTER_PROCESS_INTERVAL = "db.md.process.interval"
    
    STR_WS_NAME_AREA_MD = "db.md.area"
    STR_WS_NAME_AREA_STEP_MD = "db.md.area.step"
    
    STR_WS_NAME_PLACE_MD = "db.md.place"
    
    STR_WS_NAME_KPI = "db.md.kpi"
    STR_WS_NAME_KPI_ON_TIME_LIMITS = "db.md.kpi.on.time.limits"
    
    STR_WS_NAME_MATERIAL_GROUP_WH_RESOLVER = "db.md.mat_grp_wh.resolver"
    
    STR_WS_NAME_CONDITION_PROCESS_MATERIAL = "db.md.condition_process_mat"
End Sub
