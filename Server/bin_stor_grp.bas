Attribute VB_Name = "bin_stor_grp"
Option Explicit

Public Const STR_OUTBOUND_JIRNY As String = "JIRNY"
Public Const STR_INBOUND_HBW As String = "HBW"
Public Const STR_INBOUND_PROCESSING_PAINT As String = "PAINT"

Public Function is_outbound(str_bin_stor_grp As String)
    is_outbound = False

    If str_bin_stor_grp = STR_OUTBOUND_JIRNY Then
        is_outbound = True
    End If
End Function

Public Function is_hbw(str_bin_stor_grp As String)
    is_hbw = False

    If str_bin_stor_grp = STR_INBOUND_HBW Then
        is_hbw = True
    End If
End Function

Public Function is_processing(str_bin_stor_grp As String)
    is_processing = False

    If str_bin_stor_grp = STR_INBOUND_PROCESSING_PAINT Then
        is_processing = True
    End If
End Function
