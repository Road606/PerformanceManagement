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
Public col_process_intervals As Collection
Public col_areas As Collection
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

    Set obj_mdl_process = New MDLProcessMaster
    
    obj_mdl_process.single_data_provider.STR_WS_NAME = STR_WS_NAME_MASTER_PROCESS
    obj_mdl_process.add_listener Me
    obj_mdl_process.load_data
    
    Set obj_mdl_process_version = New MDLProcessVersionMaster
    
    obj_mdl_process_version.single_data_provider.STR_WS_NAME = STR_WS_NAME_MASTER_PROCESS_VERSION
    obj_mdl_process_version.add_listener Me
    obj_mdl_process_version.load_data
    
    Set obj_mdl_process_interval = New MDLProcessIntervalMD
    
    
    obj_mdl_process_interval.add_listener Me
    obj_mdl_process_interval.single_data_provider.STR_WS_NAME = STR_WS_NAME_MASTER_PROCESS_INTERVAL
    obj_mdl_process_interval.load_data
    
    Set obj_mdl_process = Nothing
    Set obj_mdl_process_version = Nothing
    Set obj_mdl_process_interval = Nothing
End Function

Private Function load_area_steps()
    Dim obj_mdl_area As MDLAreaMD
    Dim obj_mdl_area_step As MDLAreaStepMD

    Set obj_mdl_area = New MDLAreaMD
    obj_mdl_area.add_listener Me
    obj_mdl_area.single_data_provider.STR_WS_NAME = STR_WS_NAME_AREA_MD '"db.md.area"
    
    Set obj_mdl_area_step = New MDLAreaStepMD
    obj_mdl_area_step.add_listener Me
    obj_mdl_area_step.single_data_provider.STR_WS_NAME = STR_WS_NAME_AREA_STEP_MD '"db.md.area.step"
    
    obj_mdl_area.load_data
    obj_mdl_area_step.load_data
    
    
    Set obj_mdl_area = Nothing
    Set obj_mdl_area_step = Nothing
End Function

Private Function load_places()
    Dim obj_mdl_place As MDLPlaceMD

    Set obj_mdl_place = New MDLPlaceMD
    obj_mdl_place.add_listener Me
    obj_mdl_place.single_data_provider.STR_WS_NAME = STR_WS_NAME_PLACE_MD
    obj_mdl_place.load_data
    
    
    Set obj_mdl_place = Nothing
End Function

Public Function load_kpis()
    Dim obj_mdl_kpi As MDLKPIMD
    Dim obj_mdl_kpi_limit As MDLKPIOnTimeLimitsMD

    Set obj_mdl_kpi = New MDLKPIMD
    obj_mdl_kpi.add_listener Me
    obj_mdl_kpi.single_data_provider.STR_WS_NAME = STR_WS_NAME_KPI
    obj_mdl_kpi.load_data
    
    Set obj_mdl_kpi_limit = New MDLKPIOnTimeLimitsMD
    'obj_mdl_kpi_limit.init
    obj_mdl_kpi_limit.add_listener Me
    obj_mdl_kpi_limit.single_data_provider.STR_WS_NAME = STR_WS_NAME_KPI_ON_TIME_LIMITS
    obj_mdl_kpi_limit.load_data
    
    'Set col_kpis = obj_mdl_kpi.col_kpis
End Function

Public Function load_material_group_wh_resolvers()
    Dim obj_mdl_material_group_wh_resolver As MDLMaterialGroupWHResolver
    
    Set obj_mdl_material_group_wh_resolver = New MDLMaterialGroupWHResolver
    obj_mdl_material_group_wh_resolver.add_listener Me
    obj_mdl_material_group_wh_resolver.single_data_provider.STR_WS_NAME = STR_WS_NAME_MATERIAL_GROUP_WH_RESOLVER
    obj_mdl_material_group_wh_resolver.load_data
End Function

Public Function load_conditions_process_material()
    Dim obj_mdl_condition_process_material As MDLConditionProcessMaterial
    
    Set obj_mdl_condition_process_material = New MDLConditionProcessMaterial
    obj_mdl_condition_process_material.single_data_provider.STR_WS_NAME = STR_WS_NAME_CONDITION_PROCESS_MATERIAL
    obj_mdl_condition_process_material.add_listener Me
    obj_mdl_condition_process_material.load_data
    'Set col_conditions_process_material = obj_mdl_condition_process_material.col_vendors
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
    
    Set col_areas = New Collection
    Set col_area_steps = New Collection
    Set col_places = New Collection
    Set col_kpis = New Collection
    Set col_material_group_wh_resolvers = New Collection
    Set col_conditions_process_material = New Collection
    Set col_processes = New Collection
    Set col_process_versions = New Collection
    Set col_process_intervals = New Collection
End Sub

Public Function get_collection(mdl_id As String) As Collection
    Select Case mdl_id
        Case new_const_mdl_id.STR_MDL_AREA:
            Set get_collection = col_areas
        
        Case new_const_mdl_id.STR_MDL_PLACE:
            Set get_collection = col_places
            
        Case new_const_mdl_id.STR_MDL_AREA_STEP:
            Set get_collection = col_area_steps
        
        Case new_const_mdl_id.STR_MDL_KPI, new_const_mdl_id.STR_MDL_KPI_ONTIME_LIMITS:
            Set get_collection = col_kpis
        
        Case new_const_mdl_id.STR_MDL_PROCESS_MASTER:
            Set get_collection = col_processes
        
        Case new_const_mdl_id.STR_MDL_MATERIAL_GROUP_WH_RESOLVER:
            Set get_collection = col_material_group_wh_resolvers
            
        Case new_const_mdl_id.STR_MDL_CONDITION_PROCESS_MATERIAL:
            Set get_collection = col_conditions_process_material
        
        Case new_const_mdl_id.STR_MDL_PROCESS_INTERVAL:
            Set get_collection = col_process_intervals
        
        Case new_const_mdl_id.STR_MDL_PROCESS_VERSION_MASTER:
            Set get_collection = col_process_versions
        
    End Select
            
End Function
