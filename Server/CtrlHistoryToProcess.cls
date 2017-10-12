VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtrlHistoryToProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Public obj_history As DBHistoryRecord
'Public obj_master As ProcessMaster
'
''' main
''  ' pallet level
''Public str_pallet_id As String
''Public str_pallet_material As String
''Public str_pallet_source_destination As String
''  ' process level
''Public str_process_id As String
''Public str_process_version As String
''  ' step level
''Public str_process_step_user As String
''Public str_process_step_date_start As String
''Public str_process_step_date_end As String
''Public str_process_step_bin_from As String
''Public str_process_step_bin_to As String
''Public str_process_step_place_from As String
''Public str_process_step_place_to As String
''Public str_process_step_order_status As String
'
'Public Property Get str_pallet_id_start() As String
'    str_pallet_id_start = obj_history.str_combi_vhu_from
'End Property
'
'Public Property Get str_pallet_id_end() As String
'    str_pallet_id_end = obj_history.str_combi_vhu_to
'End Property
'
'Public Property Get str_pallet_material() As String
'    str_pallet_material = obj_history.str_material
'End Property
'
'Public Property Get str_pallet_source_destination() As String
'    str_pallet_source_destination = obj_history.str_wc_shipto
'End Property
'
'Public Property Get str_process_id() As String
'    If obj_master Is Nothing Then
'        Set obj_master = get_master
'    End If
'
'    str_process_id = obj_master.str_process_id
'End Property
'
'Public Property Get str_process_version() As String
'    If obj_master Is Nothing Then
'        Set obj_master = get_master
'    End If
'
'    str_process_version = obj_master.get_version(resolve_master_version(obj_history, obj_master)).str_id
'End Property
'
'Public Property Get str_process_step_user() As String
'    str_process_step_user = obj_history.str_user_name
'End Property
'
'Public Property Get str_process_step_date_start() As String
'    str_process_step_date_start = obj_history.str_transaction_started
'End Property
'
'Public Property Get str_process_step_date_end() As String
'    str_process_step_date_end = obj_history.str_transaction_finished
'End Property
'
'Public Property Get str_process_step_bin_from() As String
'    str_process_step_bin_from = obj_history.str_bin_from
'End Property
'
'Public Property Get str_process_step_bin_to() As String
'    str_process_step_bin_to = obj_history.str_bin_to
'End Property
'
'Public Property Get str_process_step_place_from() As String
'    str_process_step_place_from = bin_place_grp.get_place_grp(obj_history.str_bin_from)
'End Property
'
'Public Property Get str_process_step_place_to() As String
'    str_process_step_place_to = bin_place_grp.get_place_grp(obj_history.str_bin_to)
'End Property
'
'Public Property Get str_process_step_order_status() As String
'    str_process_step_order_status = "-1"
'End Property
'
'Public Property Get bool_is_single() As Boolean
'    bool_is_single = obj_history.is_single_transaction
'End Property
'
''Private Function get_master() As ProcessMaster
''    Dim obj_current_master As ProcessMaster
''    Dim obj_action As ProcessMasterAction
''    Dim obj_condition As TransactionCondition
''    Dim bool_is_found As Boolean
''
''    For Each obj_current_master In new_ctrl_process_master.col_process_masters
''        Set obj_action = obj_current_master.col_actions.Item(new_ctrl_process_master_action.STR_CREATE)
''        For Each obj_condition In obj_action.col_conditions
''            bool_is_found = obj_condition.is_match(obj_history)
''            If bool_is_found Then
''                Set get_master = obj_current_master
''                'str_process_id = obj_current_master.str_process_id
''                Exit For
''            End If
''        Next
''
''        If bool_is_found Then
''            Exit For
''        End If
''    Next
''End Function
''
''Private Function resolve_master_version(obj_record As DBHistoryRecord, obj_master As ProcessMaster) As String
''    Dim obj_version_resolver As Object
''
''    Select Case obj_master.str_version_determinant
''        Case new_ctrl_process_master_version.STR_CREATION_METHOD_SUPPLY
''            Set obj_version_resolver = New VersionOutbound
''        Case new_ctrl_process_master_version.STR_CREATION_METHOD_CREATE
''            Set obj_version_resolver = New VersionSingle
''    End Select
''
''    resolve_master_version = obj_version_resolver.retrieve(obj_record)
''End Function
