Attribute VB_Name = "bin"
Option Explicit

Public Const str_module = "BIN"

Public Const STR_USER_BIN_PREFIX = "U_"
Public STR_BIN_SEPARATOR As String

Public INT_BUILDING_PREFIX_LEN As Integer
Public STR_BUILDING_A_PREFIX As String
Public STR_BUILDING_B_PREFIX As String
Public STR_BUILDING_C_PREFIX As String
Public STR_BUILDING_HBW_IN_PREFIX As String
Public STR_BUILDING_HBW_OUT_PREFIX As String

Public INT_HALL_PREFIX_LEN As Integer
Public INT_HOUSE_PREFIX_LEN As Integer

Public STR_BUILDING_A As String
Public STR_BUILDING_B As String
Public STR_BUILDING_C As String
Public STR_BUILDING_HBW_IN As String
Public STR_BUILDING_HBW_OUT As String

Public INT_HOUSE_START As Integer
Public INT_HOUSE_LEN As Integer

Public STR_VNA1_PREFIX As String
Public STR_VNA2_PREFIX As String
Public STR_VNA3_PREFIX As String

Public STR_RA_OTHERS_PREFIX_1 As String
Public STR_RA_OTHERS_PREFIX_2 As String
Public STR_RA_OTHERS_PREFIX_3 As String

Public STR_RA_GATE_PREFIX_1 As String
Public STR_RA_GATE_PREFIX_2 As String
Public STR_RA_GATE_PREFIX_3 As String

Public STR_RA_GATE_INBOUND_PREFIX_1 As String
Public STR_RA_GATE_OUTBOUND_PREFIX_1 As String

Public INT_RA_GATE_START As Integer
Public INT_RA_GATE_LEN As Integer

'Public INT_RA_GATE_PUSH_PREFIX_LEN As Integer
'Public STR_RA1_GATE_PUSH_PREFIX As String
'Public STR_RA2_GATE_PUSH_PREFIX As String
'Public STR_RA3_GATE_PUSH_PREFIX As String

' TAs
Public INT_TA_PREFIX_LEN As Integer
Public STR_TA_PREFIX_1 As String
Public STR_TA_PREFIX_2 As String
Public STR_TA_PREFIX_3 As String
Public STR_TA_PREFIX_4 As String
Public STR_TA_PREFIX_5 As String
Public STR_TA_PREFIX_6 As String
Public STR_TA_PREFIX_7 As String
Public STR_TA_PREFIX_8 As String

Public STR_TA_HOUSE_PREFIX As String

Public INT_QUALITY_PREFIX_LEN As Integer
Public STR_QUALITY_B_PREFIX As String
Public STR_QUALITY_C_PREFIX As String

' PA
Public STR_PA_1A As String
Public STR_PA_2B As String
Public STR_PA_2C As String
Public STR_PA_3B As String

' Production
  ' LINEs
Public INT_LINE_LEN As Integer
' HALLs
Public STR_HALL_A_PREFIX As String
Public STR_HALL_B_PREFIX As String
Public STR_HALL_C_PREFIX As String

Public STR_HALL_SIDE_PREFIX As String
Public INT_HALL_SIDE_START As Integer
Public INT_HALL_SIDE_LEN As Integer

' HBW
Public INT_HBW_BUILDING_LEN As Integer
Public STR_HBW_BUILDING_A_PREFIX As String
Public STR_HBW_BUILDING_B_PREFIX As String
Public STR_HBW_BUILDING_C_PREFIX As String
Public STR_HBW_BUILDING_HBW_OUTSIDE_PREFIX As String
Public STR_HBW_BUILDING_HBW_WH_PREFIX As String

  ' gate
Public INT_HBW_GATE_START As Integer
Public INT_HBW_GATE_LEN As Integer
Public INT_HBW_GATE_LOWEST As String
Public INT_HBW_GATE_HIGHEST As String

  ' conveyor
    ' general
Public INT_HBW_CONVEYOR_START As Integer
    ' in
Public INT_HBW_CONVEYOR_IN_LEN As Integer
Public STR_HBW_CONVEYOR_IN_PREFIX_1 As String
Public STR_HBW_CONVEYOR_IN_PREFIX_2 As String
Public STR_HBW_CONVEYOR_IN_PREFIX_3 As String
    ' out
Public INT_HBW_CONVEYOR_OUT_LEN As Integer
Public STR_HBW_CONVEYOR_OUT_PREFIX_1 As String
Public STR_HBW_CONVEYOR_OUT_PREFIX_2 As String
Public STR_HBW_CONVEYOR_OUT_PREFIX_3 As String

  ' robot
Public INT_HBW_ROBOT_START As Integer
Public INT_HBW_ROBOT_DIVIDER As Integer
Public INT_HBW_ROBOT_DIVIDER_START As Integer
Public INT_HBW_ROBOT_DIVIDER_LEN As Integer
    
Public INT_HBW_ROBOT_IN_LEN As Integer
Public STR_HBW_ROBOT_IN_PREFIX_1 As String
Public STR_HBW_ROBOT_IN_PREFIX_2 As String
Public STR_HBW_ROBOT_IN_PREFIX_3 As String
    
Public INT_HBW_ROBOT_OUT_LEN As Integer
Public STR_HBW_ROBOT_OUT_PREFIX_1 As String
Public STR_HBW_ROBOT_OUT_PREFIX_2 As String
Public STR_HBW_ROBOT_OUT_PREFIX_3 As String

Public INT_HBW_LIFT_START As Integer
Public INT_HBW_LIFT_LEN As Integer
Public STR_HBW_LIFT_PREFIX As String

Public STR_HBW_BULK_PREFIX As String
Public INT_HBW_BULK_START As Integer
Public INT_HBW_BULK_LEN As Integer

Public INT_MATERIAL_HANDLING_LEN As Integer
Public STR_MATERIAL_HANDLING_SPE_PREFIX As String
Public STR_MATERIAL_HANDLING_LPE_PREFIX As String
Public STR_MATERIAL_HANDLING_ATLET_PREFIX As String

Public INT_SCALE_STATION_LEN As Integer
Public STR_SCALE_STATION_PREFIX_A As String
Public STR_SCALE_STATION_PREFIX_B As String
Public STR_SCALE_STATION_PREFIX_C As String

Public STR_BIN_DUMMY As String

Public INT_VNA_INBOUND_HOUSE_LIMIT As Integer

Public col_bin_prod_lines As Collection

Public Function init()
    STR_BIN_SEPARATOR = "-"
    
    INT_BUILDING_PREFIX_LEN = "4"
    STR_BUILDING_A_PREFIX = "6-02"
    STR_BUILDING_B_PREFIX = "6-12"
    STR_BUILDING_C_PREFIX = "6-13"
    STR_BUILDING_HBW_IN_PREFIX = "6-21"
    STR_BUILDING_HBW_OUT_PREFIX = "6-20"
    
    INT_HALL_PREFIX_LEN = "7"
    INT_HOUSE_PREFIX_LEN = "14"
    
    INT_HOUSE_START = 12
    INT_HOUSE_LEN = 3
    
    STR_BUILDING_A = "A"
    STR_BUILDING_B = "B"
    STR_BUILDING_C = "C"
    STR_BUILDING_HBW_IN = "HBW-IN"
    STR_BUILDING_HBW_OUT = "HBW-OUT"
    
    STR_VNA1_PREFIX = "6-02-01"
    STR_VNA2_PREFIX = "6-12-01"
    STR_VNA3_PREFIX = "6-13-01"
    
    STR_RA_OTHERS_PREFIX_1 = "6-02-02"
    STR_RA_OTHERS_PREFIX_2 = "6-12-02"
    STR_RA_OTHERS_PREFIX_3 = "6-13-02"
            
    STR_RA_GATE_PREFIX_1 = "6-02-80"
    STR_RA_GATE_PREFIX_2 = "6-12-80"
    STR_RA_GATE_PREFIX_3 = "6-13-80"
    
    INT_RA_GATE_START = 9
    INT_RA_GATE_LEN = 2
    STR_RA_GATE_INBOUND_PREFIX_1 = "80"
    STR_RA_GATE_OUTBOUND_PREFIX_1 = "90"
    
'    INT_RA_GATE_PUSH_PREFIX_LEN = 10
'    STR_RA1_GATE_PUSH_PREFIX = "6-02-80-90"
'    STR_RA2_GATE_PUSH_PREFIX = "6-12-80-90"
'    STR_RA3_GATE_PUSH_PREFIX = "6-13-80-90"
    
    ' TAs
     ' prefix
    INT_TA_PREFIX_LEN = 7
    STR_TA_PREFIX_1 = "6-02-11"
    STR_TA_PREFIX_2 = "6-02-22"
    STR_TA_PREFIX_3 = "6-12-33"
    STR_TA_PREFIX_4 = "6-12-44"
    STR_TA_PREFIX_5 = "6-12-55"
    STR_TA_PREFIX_6 = "6-13-66"
    STR_TA_PREFIX_7 = "6-13-77"
    STR_TA_PREFIX_8 = "6-13-88"
     ' house prefix
    STR_TA_HOUSE_PREFIX = "999"
    
    
    
    ' Quality
     ' prefix
    INT_QUALITY_PREFIX_LEN = 7
    STR_QUALITY_B_PREFIX = "6-12-05"
    STR_QUALITY_C_PREFIX = "6-13-05"
    
    ' PA
    STR_PA_1A = "6-02-20-01-999-01"
    STR_PA_2B = "6-12-20-02-999-01"
    STR_PA_2C = "6-13-20-02-999-01"
    STR_PA_3B = "6-12-20-03-999-01"
    
    ' Production
      ' line
    INT_LINE_LEN = 6
      
    STR_HALL_A_PREFIX = "6-02-03"
    STR_HALL_B_PREFIX = "6-12-03"
    STR_HALL_C_PREFIX = "6-13-03"
    
    STR_HALL_SIDE_PREFIX = "99"
    INT_HALL_SIDE_START = 12
    INT_HALL_SIDE_LEN = 2
    
    INT_HBW_BUILDING_LEN = 7
    STR_HBW_BUILDING_A_PREFIX = "6-02-20"
    STR_HBW_BUILDING_B_PREFIX = "6-12-20"
    STR_HBW_BUILDING_C_PREFIX = "6-13-20"
    STR_HBW_BUILDING_HBW_OUTSIDE_PREFIX = "6-20-02"
    STR_HBW_BUILDING_HBW_WH_PREFIX = "6-21-01"
    
    INT_HBW_CONVEYOR_START = 12
    
    INT_HBW_CONVEYOR_IN_LEN = 2
    STR_HBW_CONVEYOR_IN_PREFIX_1 = "07"
    STR_HBW_CONVEYOR_IN_PREFIX_2 = "27"
    STR_HBW_CONVEYOR_IN_PREFIX_3 = "38"
        
    INT_HBW_CONVEYOR_OUT_LEN = 2
    STR_HBW_CONVEYOR_OUT_PREFIX_1 = "01"
    STR_HBW_CONVEYOR_OUT_PREFIX_2 = "02"
    STR_HBW_CONVEYOR_OUT_PREFIX_3 = "23"
    
      ' gate
    INT_HBW_GATE_START = 13
    INT_HBW_GATE_LEN = 2
    INT_HBW_GATE_LOWEST = 1
    INT_HBW_GATE_HIGHEST = 10
    
    ' robot
    INT_HBW_ROBOT_START = 12
    INT_HBW_ROBOT_DIVIDER = 5
    INT_HBW_ROBOT_DIVIDER_START = 14
    INT_HBW_ROBOT_DIVIDER_LEN = 1
    
    INT_HBW_ROBOT_IN_LEN = 2
    STR_HBW_ROBOT_IN_PREFIX_1 = "04"
    STR_HBW_ROBOT_IN_PREFIX_2 = "24"
    STR_HBW_ROBOT_IN_PREFIX_3 = "37"
    
    INT_HBW_ROBOT_OUT_LEN = 2
    STR_HBW_ROBOT_OUT_PREFIX_1 = "06"
    STR_HBW_ROBOT_OUT_PREFIX_2 = "26"
    STR_HBW_ROBOT_OUT_PREFIX_3 = "37"
    
    INT_HBW_LIFT_START = 9
    INT_HBW_LIFT_LEN = 2
    STR_HBW_LIFT_PREFIX = "03"
    
    STR_HBW_BULK_PREFIX = "999"
    INT_HBW_BULK_START = 12
    INT_HBW_BULK_LEN = 3
    
    INT_MATERIAL_HANDLING_LEN = 3
    STR_MATERIAL_HANDLING_ATLET_PREFIX = "ATL"
    STR_MATERIAL_HANDLING_LPE_PREFIX = "LPE"
    STR_MATERIAL_HANDLING_SPE_PREFIX = "SPE"
    
    INT_SCALE_STATION_LEN = 7
    STR_SCALE_STATION_PREFIX_A = "6-02-08"
    STR_SCALE_STATION_PREFIX_B = "6-12-08"
    STR_SCALE_STATION_PREFIX_C = "6-13-08"
    
    STR_BIN_DUMMY = "999"
    
    INT_VNA_INBOUND_HOUSE_LIMIT = 800
    
    Set col_bin_prod_lines = New Collection
End Function

Public Function is_user_bin(str_bin As String) As Boolean
    is_user_bin = False
    
    If Left(str_bin, 2) = STR_USER_BIN_PREFIX Then
        is_user_bin = True
    End If
End Function

Public Function is_ra(str_bin As String) As Boolean
    is_ra = False

    Select Case Left(str_bin, INT_HALL_PREFIX_LEN)
        Case STR_RA_OTHERS_PREFIX_1, _
                STR_RA_OTHERS_PREFIX_2, _
                STR_RA_OTHERS_PREFIX_3
            is_ra = True
        Case STR_RA_GATE_PREFIX_1, _
                STR_RA_GATE_PREFIX_2, _
                STR_RA_GATE_PREFIX_3
            is_ra = True
    End Select
    
End Function

Public Function is_ra_gate_inbound(str_bin As String) As Boolean
    is_ra_gate_inbound = False

    If is_ra(str_bin) Then
        If Mid(str_bin, INT_RA_GATE_START, INT_RA_GATE_LEN) = STR_RA_GATE_INBOUND_PREFIX_1 Then
            is_ra_gate_inbound = True
        End If
    End If
End Function

Public Function is_ra_gate_outbound(str_bin As String) As Boolean
    is_ra_gate_outbound = False

    If is_ra(str_bin) Then
        If Mid(str_bin, INT_RA_GATE_START, INT_RA_GATE_LEN) = STR_RA_GATE_OUTBOUND_PREFIX_1 Then
            is_ra_gate_outbound = True
        End If
    End If
End Function

'Public Function is_gate_push(str_bin As String) As Boolean
'    is_gate_push = False
'
'    Select Case Left(str_bin, INT_RA_GATE_PUSH_PREFIX_LEN)
'        Case STR_RA1_GATE_PUSH_PREFIX, _
'                STR_RA2_GATE_PUSH_PREFIX, _
'                STR_RA3_GATE_PUSH_PREFIX
'            is_gate_push = True
'    End Select
'End Function

Public Function is_pa(str_bin As String)
    is_pa = False

    Select Case str_bin
        Case STR_PA_1A, STR_PA_2B, STR_PA_2C, STR_PA_3B
            is_pa = True
    End Select
End Function

Public Function is_production_line(str_bin As String) As Boolean
    is_production_line = False

    If Len(str_bin) = INT_LINE_LEN Then
        is_production_line = True
    End If
End Function

Public Function is_production_line_in(str_bin As String) As Boolean
    is_production_line_in = False

    If Len(str_bin) = INT_LINE_LEN Then
        is_production_line_in = True
    End If
End Function

Public Function is_production_hall(str_bin As String) As Boolean
    is_production_hall = False

    Select Case Left(str_bin, INT_HALL_PREFIX_LEN)
        Case STR_HALL_A_PREFIX, _
                STR_HALL_B_PREFIX, _
                STR_HALL_C_PREFIX
            is_production_hall = True
    End Select
End Function

Public Function is_production_hall_side(str_bin As String) As Boolean
    is_production_hall_side = False

    Select Case Left(str_bin, INT_HALL_PREFIX_LEN)
        Case STR_HALL_A_PREFIX, _
                STR_HALL_B_PREFIX, _
                STR_HALL_C_PREFIX
            If Mid(str_bin, INT_HALL_SIDE_START, INT_HALL_SIDE_LEN) = STR_HALL_SIDE_PREFIX Then
                is_production_hall_side = True
            End If
    End Select
End Function

Public Function is_vna_bulk(str_bin As String) As Boolean
    is_vna_bulk = False

    Select Case Left(str_bin, INT_HALL_PREFIX_LEN)
        Case STR_VNA1_PREFIX, _
                STR_VNA2_PREFIX, _
                STR_VNA3_PREFIX
            If CInt(Right(Left(str_bin, INT_HOUSE_PREFIX_LEN), INT_HOUSE_LEN)) > INT_VNA_INBOUND_HOUSE_LIMIT Then
                is_vna_bulk = True
            End If
    End Select
End Function

Public Function is_vna_rack(str_bin As String) As Boolean
    is_vna_rack = False

    Select Case Left(str_bin, INT_HALL_PREFIX_LEN)
        Case STR_VNA1_PREFIX, _
                STR_VNA2_PREFIX, _
                STR_VNA3_PREFIX
            If CInt(Right(Left(str_bin, INT_HOUSE_PREFIX_LEN), INT_HOUSE_LEN)) < INT_VNA_INBOUND_HOUSE_LIMIT Then
                is_vna_rack = True
            End If
    End Select
End Function

Public Function is_ta_bulk(str_bin As String) As Boolean
    is_ta_bulk = False
    
    Select Case Left(str_bin, INT_TA_PREFIX_LEN)
        Case STR_TA_PREFIX_1, _
                STR_TA_PREFIX_2, _
                STR_TA_PREFIX_3, _
                STR_TA_PREFIX_4, _
                STR_TA_PREFIX_5, _
                STR_TA_PREFIX_6, _
                STR_TA_PREFIX_7, _
                STR_TA_PREFIX_8
            If Mid(str_bin, INT_HOUSE_START, INT_HOUSE_LEN) = STR_TA_HOUSE_PREFIX Then
                is_ta_bulk = True
            End If
    End Select
End Function

Public Function is_ta_rack(str_bin As String) As Boolean
    Select Case Left(str_bin, INT_TA_PREFIX_LEN)
        Case STR_TA_PREFIX_1, _
                STR_TA_PREFIX_2, _
                STR_TA_PREFIX_3, _
                STR_TA_PREFIX_4, _
                STR_TA_PREFIX_5, _
                STR_TA_PREFIX_6, _
                STR_TA_PREFIX_7, _
                STR_TA_PREFIX_8
            If Mid(str_bin, INT_HOUSE_START, INT_HOUSE_LEN) <> STR_TA_HOUSE_PREFIX Then
                is_ta_rack = True
            End If
    End Select
End Function

Public Function is_hbw(str_bin As String) As Boolean
    is_hbw = False
    
    Select Case Left(str_bin, INT_HBW_BUILDING_LEN)
        Case STR_HBW_BUILDING_A_PREFIX, _
                STR_HBW_BUILDING_B_PREFIX, _
                STR_HBW_BUILDING_C_PREFIX, _
                STR_HBW_BUILDING_HBW_OUTSIDE_PREFIX, _
                STR_HBW_BUILDING_HBW_WH_PREFIX
            is_hbw = True
    End Select
End Function

Public Function is_hbw_conveyor_in(str_bin As String) As Boolean
    is_hbw_conveyor_in = False

    If is_hbw(str_bin) Then
        If Mid(str_bin, INT_HBW_CONVEYOR_START, INT_HBW_CONVEYOR_IN_LEN) = STR_HBW_CONVEYOR_IN_PREFIX_1 Or _
            Mid(str_bin, INT_HBW_CONVEYOR_START, INT_HBW_CONVEYOR_IN_LEN) = STR_HBW_CONVEYOR_IN_PREFIX_2 Or _
            Mid(str_bin, INT_HBW_CONVEYOR_START, INT_HBW_CONVEYOR_IN_LEN) = STR_HBW_CONVEYOR_IN_PREFIX_3 _
                Then
            is_hbw_conveyor_in = True
        End If
    End If
End Function

Public Function is_hbw_conveyor_out(str_bin As String) As Boolean
    is_hbw_conveyor_out = False

    If is_hbw(str_bin) Then
        If Mid(str_bin, INT_HBW_CONVEYOR_START, INT_HBW_CONVEYOR_OUT_LEN) = STR_HBW_CONVEYOR_OUT_PREFIX_1 Or _
            Mid(str_bin, INT_HBW_CONVEYOR_START, INT_HBW_CONVEYOR_OUT_LEN) = STR_HBW_CONVEYOR_OUT_PREFIX_2 Or _
            Mid(str_bin, INT_HBW_CONVEYOR_START, INT_HBW_CONVEYOR_OUT_LEN) = STR_HBW_CONVEYOR_OUT_PREFIX_3 _
                Then
            is_hbw_conveyor_out = True
        End If
    End If
End Function

Public Function is_hbw_gate(str_bin As String) As Boolean
    Dim int_divider As Integer
    
    is_hbw_gate = False

    If is_hbw(str_bin) Then
        int_divider = CInt(Mid(str_bin, INT_HBW_GATE_START, INT_HBW_GATE_LEN))
        
        If int_divider > INT_HBW_GATE_LOWEST And int_divider < INT_HBW_GATE_HIGHEST Then
            is_hbw_gate = True
        End If
    End If
End Function

Public Function is_hbw_robot_in(str_bin As String) As Boolean
    Dim str_prefix As String
    
    is_hbw_robot_in = False

    If is_hbw(str_bin) Then
        str_prefix = Mid(str_bin, INT_HBW_ROBOT_START, INT_HBW_ROBOT_IN_LEN)
        
        If str_prefix = STR_HBW_ROBOT_IN_PREFIX_1 Or str_prefix = STR_HBW_ROBOT_IN_PREFIX_2 Then
            is_hbw_robot_in = True
        ElseIf str_prefix = STR_HBW_ROBOT_IN_PREFIX_3 And _
                CInt(Mid(str_bin, INT_HBW_ROBOT_DIVIDER_START, INT_HBW_ROBOT_DIVIDER_LEN)) < INT_HBW_ROBOT_DIVIDER Then
            is_hbw_robot_in = True
        End If
    End If
End Function

Public Function is_hbw_robot_out(str_bin As String) As Boolean
    Dim str_prefix As String

    is_hbw_robot_out = False

    If is_hbw(str_bin) Then
        str_prefix = Mid(str_bin, INT_HBW_ROBOT_START, INT_HBW_ROBOT_OUT_LEN)
        
        If str_prefix = STR_HBW_ROBOT_OUT_PREFIX_1 Or str_prefix = STR_HBW_ROBOT_OUT_PREFIX_2 Then
            is_hbw_robot_out = True
        ElseIf str_prefix = STR_HBW_ROBOT_OUT_PREFIX_3 And _
                CInt(Mid(str_bin, INT_HBW_ROBOT_DIVIDER_START, INT_HBW_ROBOT_DIVIDER_LEN)) > INT_HBW_ROBOT_DIVIDER Then
            is_hbw_robot_out = True
        End If
    End If
End Function

Public Function is_hbw_wh(str_bin As String) As Boolean
    is_hbw_wh = False
    
    If Left(str_bin, INT_HBW_BUILDING_LEN) = STR_HBW_BUILDING_HBW_WH_PREFIX Then
        is_hbw_wh = True
    End If
End Function

Public Function is_hbw_lift(str_bin As String) As Boolean
    Dim str_prefix As String

    is_hbw_lift = False

    If is_hbw(str_bin) Then
        str_prefix = Mid(str_bin, INT_HBW_LIFT_START, INT_HBW_LIFT_LEN)
        
        If str_prefix = STR_HBW_LIFT_PREFIX Then
            is_hbw_lift = True
        End If
    End If
End Function

Public Function is_material_handling(str_bin) As Boolean
    is_material_handling = False
        
    Select Case Left(str_bin, INT_MATERIAL_HANDLING_LEN)
        Case STR_MATERIAL_HANDLING_ATLET_PREFIX, _
                STR_MATERIAL_HANDLING_LPE_PREFIX, _
                STR_MATERIAL_HANDLING_SPE_PREFIX
            is_material_handling = True
    End Select
End Function

Public Function is_scale_station(str_bin) As Boolean
    is_scale_station = False
        
    Select Case Left(str_bin, INT_SCALE_STATION_LEN)
        Case STR_SCALE_STATION_PREFIX_A, _
                STR_SCALE_STATION_PREFIX_B, _
                STR_SCALE_STATION_PREFIX_C
            is_scale_station = True
    End Select
End Function

Public Function is_quality(str_bin) As Boolean
    is_quality = False
        
    Select Case Left(str_bin, INT_QUALITY_PREFIX_LEN)
        Case STR_QUALITY_B_PREFIX, _
                STR_QUALITY_C_PREFIX
            is_quality = True
    End Select
End Function

Public Function is_dummy(str_bin) As Boolean
    is_dummy = str_bin = STR_BIN_DUMMY
End Function

Public Function get_building(str_bin As String) As String
    ' # implement - temporary solution, implement from table
    If is_production_line_in(str_bin) Then
        get_building = col_bin_prod_lines(str_bin).str_building
'        Select Case str_bin
'            Case "113339", "118931", "144324", "162013", "172935", "113333", "130616", "149870", "153585", "154147"
'                get_building = "A"
'            Case "113370", "118207", "128210", "143706", "149538", "154107", "125132", "125134", "125137", "201383", "201384", "175407", "175408"
'                get_building = "B"
'            Case "157716", "195739", "195740", "178460", "178461"
'                get_building = "C"
'        End Select
    Else
        Select Case Left(str_bin, INT_BUILDING_PREFIX_LEN)
            Case STR_BUILDING_A_PREFIX
                get_building = STR_BUILDING_A
            Case STR_BUILDING_B_PREFIX
                get_building = STR_BUILDING_B
            Case STR_BUILDING_C_PREFIX
                get_building = STR_BUILDING_C
            Case STR_BUILDING_HBW_IN_PREFIX
                get_building = STR_BUILDING_HBW_IN
            Case STR_BUILDING_HBW_OUT_PREFIX
                get_building = STR_BUILDING_HBW_OUT
        End Select
    End If
End Function

'Public Function get_gate(str_bin As String) As String
'
'    If is_gate_in(str_bin) Or is_gate_push(str_bin) Then
'        get_gate = Mid(str_bin, INT_HOUSE_START, INT_HOUSE_LEN)
'    End If
'End Function

'Public Function get_hall(str_bin As String) As String
'    get_hall = get_building(str_bin)
'
'    If is_production_hall(str_bin) Then
'        get_hall = get_production_hall(str_bin)
'    ElseIf is_ra(str_bin) Then
'        get_hall = "RA-" & get_hall
'    ElseIf is_vna(str_bin) Then
'        get_hall = "VNA-" & get_hall
'    End If
'End Function

'Public Function get_production_hall(str_bin As String) As String
'    If is_production_hall(str_bin) Then
'        Select Case Left(str_bin, INT_HALL_PREFIX_LEN)
'            Case STR_HALL_A_PREFIX
'                get_production_hall = STR_BUILDING_A
'            Case STR_HALL_B_PREFIX
'                get_production_hall = STR_BUILDING_B
'            Case STR_HALL_C_PREFIX
'                get_production_hall = STR_BUILDING_C
'        End Select
'
'        If get_production_hall <> "" Then
'            get_production_hall = get_production_hall & CInt(Mid(str_bin, INT_AISLE_START, INT_AISLE_LEN))
'        End If
'    End If
'End Function

