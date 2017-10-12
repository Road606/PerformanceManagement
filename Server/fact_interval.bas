Attribute VB_Name = "fact_interval"
Option Explicit

Public Const INT_START_LEFT As Integer = 1
Public Const INT_START_LIMIT As Integer = 2
Public Const INT_START_RIGHT As Integer = 4

Public Const INT_END_LEFT As Integer = 8
Public Const INT_END_LIMIT As Integer = 16
Public Const INT_END_RIGHT As Integer = 32

Public Const INT_START_END_LIMIT As Integer = 64
Public Const INT_START_END_RIGHT As Integer = 128

Public Const INT_END_START_LEFT As Integer = 256
Public Const INT_END_START_LIMIT As Integer = 512

Public Function create_intervals_hourly(obj_period_start As Date, obj_period_end As Date, int_minute_start As Integer, int_minute_end As Integer, int_offset_hour As Integer) As Collection
    Dim int_current_hour As Integer
    Dim obj_current_start As Date
    Dim obj_period_interval As Interval
    Dim obj_interval As Interval
    
    int_current_hour = Hour(obj_period_start)
    obj_current_start = obj_period_start
    
    Set create_intervals_hourly = New Collection
    Set obj_period_interval = New Interval
    obj_period_interval.obj_start = obj_current_start
    obj_period_interval.obj_end = obj_period_end
    
    Do While obj_current_start < obj_period_end
        Set obj_interval = New Interval
        obj_interval.obj_start = DateSerial(Year(obj_current_start), Month(obj_current_start), Day(obj_current_start))
        obj_interval.obj_start = obj_interval.obj_start + TimeSerial(Hour(obj_current_start), int_minute_start, 0)
        
        obj_interval.obj_end = DateSerial(Year(obj_current_start), Month(obj_current_start), Day(obj_current_start))
        obj_interval.obj_end = obj_interval.obj_end + TimeSerial(Hour(obj_current_start), int_minute_end, 0)
        
        If Not obj_period_interval.is_out_left(obj_interval.obj_start, obj_interval.obj_end) Then
            create_intervals_hourly.add obj_interval
        End If
                
        int_current_hour = Hour(obj_current_start) + 1
        obj_current_start = DateSerial(Year(obj_current_start), Month(obj_current_start), Day(obj_current_start))
        obj_current_start = obj_current_start + TimeSerial(int_current_hour, 0, 0)
    Loop
End Function
