Attribute VB_Name = "test_interval"
Option Explicit

Public Function test()
    Dim obj_interval As Interval
    
    Set obj_interval = New Interval
    obj_interval.obj_start = DateValue("24.8.17 10:00") + TimeValue("24.8.17 10:00")
    obj_interval.obj_end = DateValue("24.8.17 11:00") + TimeValue("24.8.17 11:00")
    
    obj_interval.calculate_status "24.8.17 10:15", "24.8.17 10:50"
    Debug.Print obj_interval.int_status
    
    Debug.Print "Is In"
    Debug.Print "True=" & obj_interval.is_in("24.8.17 10:15", "24.8.17 10:50")
    Debug.Print "False=" & obj_interval.is_in("24.8.17 09:15", "24.8.17 10:50")
    Debug.Print "False=" & obj_interval.is_in("24.8.17 10:15", "24.8.17 11:50")
    Debug.Print "False=" & obj_interval.is_in("24.8.17 09:15", "24.8.17 11:50")
    Debug.Print "False=" & obj_interval.is_in("24.8.17 09:15", "24.8.17 09:50")
    Debug.Print "False=" & obj_interval.is_in("24.8.17 11:15", "24.8.17 11:50")
    
    Debug.Print "Out Left"
    Debug.Print "False=" & obj_interval.is_out_left("24.8.17 10:15", "24.8.17 10:50")
    Debug.Print "False=" & obj_interval.is_out_left("24.8.17 09:15", "24.8.17 10:50")
    Debug.Print "False=" & obj_interval.is_out_left("24.8.17 10:15", "24.8.17 11:50")
    Debug.Print "False=" & obj_interval.is_out_left("24.8.17 09:15", "24.8.17 11:50")
    Debug.Print "True=" & obj_interval.is_out_left("24.8.17 09:15", "24.8.17 09:50")
    Debug.Print "False=" & obj_interval.is_out_left("24.8.17 11:15", "24.8.17 11:50")
    Debug.Print "True=" & obj_interval.is_out_left("24.8.17 09:15", "24.8.17 10:00")
    
    Debug.Print "Out Right"
    Debug.Print "False=" & obj_interval.is_out_right("24.8.17 10:15", "24.8.17 10:50")
    Debug.Print "False=" & obj_interval.is_out_right("24.8.17 09:15", "24.8.17 10:50")
    Debug.Print "False=" & obj_interval.is_out_right("24.8.17 10:15", "24.8.17 11:50")
    Debug.Print "False=" & obj_interval.is_out_right("24.8.17 09:15", "24.8.17 11:50")
    Debug.Print "False=" & obj_interval.is_out_right("24.8.17 09:15", "24.8.17 09:50")
    Debug.Print "True=" & obj_interval.is_out_right("24.8.17 11:15", "24.8.17 11:50")
    Debug.Print "True=" & obj_interval.is_out_right("24.8.17 11:00", "24.8.17 11:50")
    Debug.Print "False=" & obj_interval.is_out_right("24.8.17 09:15", "24.8.17 10:00")
End Function

Public Function test_fact()
    Dim obj_period_start As Date
    Dim obj_period_end As Date
    
'    obj_period_start = DateValue("24.8.17 10:00") + TimeValue("24.8.17 10:00")
'    obj_period_end = DateValue("24.8.17 11:00:00") + TimeValue("24.8.17 11:00:00")
'    fact_interval.create_intervals_hourly obj_period_start, obj_period_end, 20, 50, 20
    
    obj_period_start = DateValue("24.8.17 23:00") + TimeValue("24.8.17 23:00")
    obj_period_end = DateValue("25.8.17 01:00") + TimeValue("25.8.17 01:00")
    fact_interval.create_intervals_hourly obj_period_start, obj_period_end, 20, 50, 20
End Function
