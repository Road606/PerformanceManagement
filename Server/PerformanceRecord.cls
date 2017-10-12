VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_process_subtype As String
Public str_process_type As String
Public str_pallet_id As String
Public str_material As String
Public str_start_datetime As String
Public str_end_datetime As String
'Public lng_duration As Long
Public str_start_bin As String
Public str_end_bin As String
Public str_start_building As String
Public str_end_building As String
Public str_start_hall As String
Public str_end_hall As String

'Public str_start_date As String
'Public str_start_hour As String
'Public str_process_step As String
Public str_process_part As String
Public int_process_steps As Integer
Public str_status As String

Public str_quality_check As String

Const INT_STATUS_OPEN As Integer = 1
Const INT_STATUS_CLOSED As Integer = 2
Public int_status As Integer ' open, closed

'Public obj_parent As PerformanceInbound

Public Function open_obj()
    int_status = INT_STATUS_OPEN
End Function

Public Function close_obj()
    int_status = INT_STATUS_CLOSED
    'int_process_steps = obj_parent.int_process_steps
End Function

Public Function is_open() As Boolean
    is_open = (int_status = INT_STATUS_OPEN)
End Function

Public Function is_closed() As Boolean
    is_closed = (int_status = INT_STATUS_CLOSED)
End Function

Public Property Get lng_duration() As Long
    On Error GoTo ERR_MISSING_DATA
    lng_duration = DateDiff(hndl_process.STR_DURATION_UNIT, str_start_datetime, str_end_datetime, vbMonday, vbFirstFourDays)
    On Error GoTo 0
    Exit Function
ERR_MISSING_DATA:
    lng_duration = 0
End Property
