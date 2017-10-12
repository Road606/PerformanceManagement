VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Interval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public obj_start As Date
Public obj_end As Date

Public int_status As Integer

Public Function calculate_status(str_start As String, str_end As String)
    Dim obj_period_start As Date
    Dim obj_period_end As Date
    
    obj_period_start = DateValue(str_start) + TimeValue(str_start)
    obj_period_end = DateValue(str_end) + TimeValue(str_end)
    
    If obj_period_start < obj_start Then
        int_status = fact_interval.INT_START_LEFT
    ElseIf obj_period_start > obj_start Then
        If obj_period_start > obj_end Then
            int_status = fact_interval.INT_START_END_RIGHT
        ElseIf obj_period_start = obj_end Then
            int_status = fact_interval.INT_START_END_LIMIT
        Else
            int_status = fact_interval.INT_START_RIGHT
        End If
    Else
        int_status = fact_interval.INT_START_LIMIT
    End If
    
    
    
    If obj_period_end < obj_end Then
        If obj_period_end < obj_start Then
            int_status = int_status + fact_interval.INT_END_START_LEFT
        ElseIf obj_period_end = obj_start Then
            int_status = int_status + fact_interval.INT_END_START_LIMIT
        Else
            int_status = int_status + fact_interval.INT_END_LEFT
        End If
    ElseIf obj_period_end > obj_end Then
        int_status = int_status + fact_interval.INT_END_RIGHT
    Else
        int_status = int_status + fact_interval.INT_END_LIMIT
    End If
End Function

Public Function is_in(str_start As String, str_end As String) As Boolean
    calculate_status str_start, str_end
    
    is_in = int_status And _
        (fact_interval.INT_START_LIMIT + fact_interval.INT_START_RIGHT)
     
    If is_in Then
        is_in = int_status And (fact_interval.INT_END_LEFT + fact_interval.INT_END_LIMIT)
    End If
End Function

Public Function is_out_left(str_start As String, str_end As String) As Boolean
    calculate_status str_start, str_end
    
    is_out_left = int_status And fact_interval.INT_START_LEFT
     
    If is_out_left Then
        is_out_left = int_status And (fact_interval.INT_END_START_LEFT + fact_interval.INT_END_START_LIMIT)
    End If
End Function

Public Function is_out_right(str_start As String, str_end As String) As Boolean
    calculate_status str_start, str_end
    
    is_out_right = int_status And _
        (fact_interval.INT_START_END_LIMIT + fact_interval.INT_START_END_RIGHT)
     
    If is_out_right Then
        is_out_right = int_status And (fact_interval.INT_END_RIGHT)
    End If
End Function
