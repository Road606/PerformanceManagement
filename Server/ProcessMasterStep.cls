VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessMasterStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private private_obj_version As ProcessMasterVersion

Public str_place_from As String
Public str_place_to As String

Public col_orders As Collection

Public Property Get obj_version() As ProcessMasterVersion
    Set obj_version = private_obj_version
End Property

Public Property Let obj_version(obj_version As ProcessMasterVersion)
    Set private_obj_version = obj_version
    obj_version.add_step Me
End Property

Public Function add_order(str_order As String)
    col_orders.add str_order, str_order
End Function

Public Function is_valid(str_order As String) As Boolean
    is_valid = True

    On Error GoTo INFO_ORDER_NOT_FOUND
    col_orders.Item str_order
    On Error GoTo 0
    Exit Function
INFO_ORDER_NOT_FOUND:
    is_valid = False
End Function

Private Sub Class_Initialize()
    Set col_orders = New Collection
End Sub
