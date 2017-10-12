Attribute VB_Name = "wc"
Option Explicit

Public INT_BUILDING_PREFIX_LEN As Integer
Public STR_BUILDING_A_PREFIX As String
Public STR_BUILDING_B_PREFIX As String
Public STR_BUILDING_C_PREFIX As String

Public STR_BUILDING_A As String
Public STR_BUILDING_B As String
Public STR_BUILDING_C As String

Public INT_HALL_PREFIX_START As Integer
Public INT_HALL_PREFIX_LEN As Integer

Public STR_HALL_B2_ORIGINAL As String
Public STR_HALL_B2_NEW As String
Public STR_HALL_B4_ORIGINAL As String
Public STR_HALL_B4_NEW As String

Public STR_MODULE_FG As String
Public STR_MODULE_PP As String
Public STR_MODULE_PROC As String

Public Function init()
    INT_BUILDING_PREFIX_LEN = 1
    
    STR_BUILDING_A_PREFIX = "5"
    STR_BUILDING_B_PREFIX = "6"
    STR_BUILDING_C_PREFIX = "7"
    
    STR_BUILDING_A = "A"
    STR_BUILDING_B = "B"
    STR_BUILDING_C = "C"
        
    INT_HALL_PREFIX_START = 2
    INT_HALL_PREFIX_LEN = 1
    
    STR_HALL_B2_ORIGINAL = "4"
    STR_HALL_B2_NEW = "2"
    STR_HALL_B4_ORIGINAL = "2"
    STR_HALL_B4_NEW = "4"
    
    STR_MODULE_FG = "FG"
    STR_MODULE_PP = "PP"
    STR_MODULE_PROC = "PROC"
End Function

Public Function get_building(str_wc As String) As String
    Select Case Left(str_wc, INT_BUILDING_PREFIX_LEN)
        Case STR_BUILDING_A_PREFIX
            get_building = STR_BUILDING_A
        Case STR_BUILDING_B_PREFIX
            get_building = STR_BUILDING_B
        Case STR_BUILDING_C_PREFIX
            get_building = STR_BUILDING_C
    End Select
End Function

Public Function get_production_hall(str_wc As String) As String
    Dim str_prod_hall As String

    Select Case Left(str_wc, INT_BUILDING_PREFIX_LEN)
        Case STR_BUILDING_A_PREFIX
            get_production_hall = STR_BUILDING_A
        Case STR_BUILDING_B_PREFIX
            get_production_hall = STR_BUILDING_B
        Case STR_BUILDING_C_PREFIX
            get_production_hall = STR_BUILDING_C
    End Select
            
    If get_production_hall <> "" Then
        str_prod_hall = Mid(str_wc, INT_HALL_PREFIX_START, INT_HALL_PREFIX_LEN)
        
        If get_production_hall = STR_BUILDING_B Then
            If str_prod_hall = STR_HALL_B2_ORIGINAL Then
                str_prod_hall = STR_HALL_B2_NEW
            ElseIf str_prod_hall = STR_HALL_B4_ORIGINAL Then
                str_prod_hall = STR_HALL_B4_NEW
            End If
        End If
        
        get_production_hall = get_production_hall & str_prod_hall
    End If
End Function

Public Function get_module(str_wc As String)
    Select Case str_wc
        Case "521", "551", "552", "553", "621", "622", _
                "723", "732", "734"
            get_module = STR_MODULE_FG
        Case "650", "651", "652", "653", "654", "655"
            get_module = STR_MODULE_PROC
        Case Else
            get_module = STR_MODULE_PP
    End Select
End Function
