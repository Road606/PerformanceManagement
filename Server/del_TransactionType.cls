VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "del_TransactionType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Public str_id As String
'
'Public private_obj_master As ProcessMaster
'Public col_conditions As Collection
'
'Public Property Get obj_master() As ProcessMaster
'    Set obj_master = private_obj_master
'End Property
'
'Public Property Let obj_master(obj_master As ProcessMaster)
'    Set private_obj_master = obj_master
'    obj_master.add_action Me
'End Property
'
'Public Function add_condition(obj_condition As TransactionCondition)
'    col_conditions.add obj_condition
'End Function
'
'Public Function retrieve_condition_from_record(obj_record As DBHistoryToProcessRecord) As TransactionCondition
'    Dim obj_condition As TransactionCondition
'
'    For Each obj_condition In col_conditions
'        If obj_condition.is_match(obj_record) Then
'            Set retrieve_condition_from_record = obj_condition
'            Exit For
'        End If
'    Next
'End Function
'
'Private Sub Class_Initialize()
'    Set col_conditions = New Collection
'End Sub
