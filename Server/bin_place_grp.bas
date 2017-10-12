Attribute VB_Name = "bin_place_grp"
Option Explicit

Public Const str_module = "bin_place_grp"

Public Const STR_DUMMY As String = "DUMMY"
Public Const STR_HBW As String = "HBW"
Public Const STR_HBW_CONVEYOR_IN As String = "HBW_CONVEYOR_IN"
Public Const STR_HBW_CONVEYOR_OUT As String = "HBW_CONVEYOR_OUT"
Public Const STR_HBW_GATE As String = "HBW_GATE"
Public Const STR_HBW_LIFT As String = "HBW_LIFT"
Public Const STR_HBW_ROBOT_IN As String = "HBW_ROBOT_IN"
Public Const STR_HBW_ROBOT_OUT As String = "HBW_ROBOT_OUT"
Public Const STR_HBW_WH As String = "HBW_WH"
Public Const STR_HBW_OTHERS As String = "HBW_OTHERS"
Public Const STR_MATERIAL_HANDLING As String = "MATERIAL_HANDLING"
Public Const STR_PA As String = "PA"
Public Const str_prod_hall As String = "PROD_HALL"
Public Const STR_PROD_LINE_IN As String = "PROD_LINE_IN"
Public Const STR_PROD_LINE_OUT As String = "PROD_LINE_OUT"
Public Const STR_RA_OTHERS As String = "RA_OTHERS"
Public Const STR_RA_GATE_IN As String = "RA_BULK"
Public Const STR_RA_GATE_OUT As String = "RA_OUTBOUND"
Public Const STR_QUALITY As String = "QUALITY"
Public Const STR_SCALE_STATION As String = "SCALE_STATION"
Public Const STR_TA_BULK As String = "TA_BULK"
Public Const STR_TA_RACK As String = "TA_RACK"
Public Const STR_VNA_BULK As String = "VNA_BULK"
Public Const STR_VNA_RACK As String = "VNA_RACK"

Public Const STR_USER_BIN As String = "USERBIN"

Public Function get_place_grp(str_bin As String) As String
    If bin.is_user_bin(str_bin) Then
        get_place_grp = bin_place_grp.STR_USER_BIN
    Else
        If bin.is_vna_rack(str_bin) Then
            get_place_grp = bin_place_grp.STR_VNA_RACK
        ElseIf bin.is_vna_bulk(str_bin) Then
            get_place_grp = bin_place_grp.STR_VNA_BULK
        ElseIf bin.is_production_hall_side(str_bin) Then
            get_place_grp = bin_place_grp.str_prod_hall
        ElseIf bin.is_production_hall(str_bin) Then
            get_place_grp = bin_place_grp.STR_PROD_LINE_OUT
        ElseIf bin.is_production_line_in(str_bin) And Not bin.is_material_handling(str_bin) Then
            get_place_grp = bin_place_grp.STR_PROD_LINE_IN
        ElseIf bin.is_ta_bulk(str_bin) Then
            get_place_grp = bin_place_grp.STR_TA_BULK
        ElseIf bin.is_ta_rack(str_bin) Then
            get_place_grp = bin_place_grp.STR_TA_RACK
        ElseIf bin.is_hbw_conveyor_in(str_bin) Then
            get_place_grp = bin_place_grp.STR_HBW_CONVEYOR_IN
        ElseIf bin.is_hbw_conveyor_out(str_bin) Then
            get_place_grp = bin_place_grp.STR_HBW_CONVEYOR_OUT
        ElseIf bin.is_hbw_gate(str_bin) Then
            get_place_grp = bin_place_grp.STR_HBW_GATE
        ElseIf bin.is_hbw_robot_in(str_bin) Then
            get_place_grp = bin_place_grp.STR_HBW_ROBOT_IN
        ElseIf bin.is_hbw_robot_out(str_bin) Then
            get_place_grp = bin_place_grp.STR_HBW_ROBOT_OUT
        ElseIf bin.is_hbw_wh(str_bin) Then
            get_place_grp = bin_place_grp.STR_HBW_WH
        ElseIf bin.is_hbw_lift(str_bin) Then
            get_place_grp = bin_place_grp.STR_HBW_LIFT
        ElseIf bin.is_hbw(str_bin) Then
            get_place_grp = bin_place_grp.STR_HBW_OTHERS
        ElseIf bin.is_material_handling(str_bin) Then
            get_place_grp = bin_place_grp.STR_MATERIAL_HANDLING
        ElseIf bin.is_scale_station(str_bin) Then
            get_place_grp = bin_place_grp.STR_SCALE_STATION
        ElseIf bin.is_dummy(str_bin) Then
            get_place_grp = bin_place_grp.STR_DUMMY
        ElseIf bin.is_quality(str_bin) Then
            get_place_grp = bin_place_grp.STR_QUALITY
        ElseIf bin.is_pa(str_bin) Then
            get_place_grp = bin_place_grp.STR_PA
        ElseIf bin.is_ra_gate_inbound(str_bin) Then
            get_place_grp = bin_place_grp.STR_RA_GATE_IN
        ElseIf bin.is_ra_gate_outbound(str_bin) Then
            get_place_grp = bin_place_grp.STR_RA_GATE_OUT
        ElseIf bin.is_ra(str_bin) Then
            get_place_grp = bin_place_grp.STR_RA_OTHERS
        Else
            hndl_log.log db_log.TYPE_WARN, str_module, "get_place_grp", "Not found place group for BIN: " & str_bin
        End If
    End If
End Function
