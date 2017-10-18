VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KPIOnTimeInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public str_id As String

Public str_date_time As String
Public str_shift As String
Public str_shift_type As String
'Public str_date As String
'Public str_time As String
'Public col_pallets As Collection
Public col_pallets_step As Collection
Public col_pallets_full_live As Collection

Public bool_is_finished As Boolean

Public obj_next_interval As KPIOnTimeInterval
Public obj_previous_interval As KPIOnTimeInterval

Public private_obj_interval As Interval

Public Property Get str_id() As String
    str_id = str_date_time
End Property

Public Property Get obj_interval() As Interval
    Set obj_interval = private_obj_interval
End Property

Public Property Let obj_interval(obj_interval As Interval)
    Set private_obj_interval = obj_interval
    'str_id = obj_interval.obj_start
End Property

Public Property Get str_date() As String
    str_date = DateValue(str_date_time)
End Property

Public Property Get str_time() As String
    str_time = Format(TimeValue(str_date_time), "hh:mm") ' # implement format in config file
End Property
Private Sub Class_Initialize()
    Set col_pallets_step = New Collection
    Set col_pallets_full_live = New Collection
End Sub

Public Function is_valid_interval(str_date_start As String, str_date_end As String) As Boolean
    is_valid_interval = False

    private_obj_interval.calculate_status str_date_start, str_date_end
    
    If Not private_obj_interval.int_status And fact_interval.INT_START_END_RIGHT Then
        If Not private_obj_interval.int_status And fact_interval.INT_END_START_LEFT Then
            is_valid_interval = True
        End If
    End If
End Function

