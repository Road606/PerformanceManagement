VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransactionCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_type_start As String
Public str_type_end As String
Public str_code As String
Public str_user As String
Public str_task_list_type As String
Public str_place_from As String
Public str_place_to As String

Private private_obj_action As ProcessMasterAction

Public Property Get obj_action() As ProcessMasterAction
    Set obj_action = private_obj_action
End Property

Public Property Let obj_action(obj_action As ProcessMasterAction)
    Set private_obj_action = obj_action
    obj_action.add_condition Me
End Property

Public Function is_match(obj_record As DBHistoryRecord) As Boolean
    Dim obj_regex As RegExp
    
    Set obj_regex = New RegExp
    With obj_regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With
    
    ' check transaction type
      ' start
    is_match = is_match_single(obj_regex, str_type_start, obj_record.str_transaction_type_started)
      ' end
    If Not obj_record.is_single_transaction Then
        If is_match Then
            is_match = is_match_single(obj_regex, str_type_end, obj_record.str_transaction_type_finished)
        End If
    End If
    ' check transaction code
    If is_match Then
        is_match = is_match_single(obj_regex, str_code, obj_record.str_transaction_code)
    End If
    ' check user
    If is_match Then
        is_match = is_match_single(obj_regex, str_user, obj_record.str_user_name)
    End If
    ' check task list type
    If is_match Then
        is_match = is_match_single(obj_regex, str_task_list_type, obj_record.str_task_list_type) 'obj_record.str_additional_task_list_type)
    End If
    ' check place from
    If is_match Then
        is_match = is_match_single(obj_regex, str_place_from, bin_place_grp.get_place_grp(obj_record.str_bin_from)) 'obj_record.str_process_step_place_from)
    End If
    ' check place to
    If is_match Then
        is_match = is_match_single(obj_regex, str_place_to, bin_place_grp.get_place_grp(obj_record.str_bin_to))
    End If
End Function

Private Function is_match_single(obj_regex As RegExp, str_pattern As String, str_checked_value As String) As Boolean
    If Len(str_pattern) = 0 Then
        is_match_single = Len(str_checked_value) = 0
    Else
        obj_regex.Pattern = str_pattern
        is_match_single = obj_regex.test(str_checked_value)
    End If
End Function
