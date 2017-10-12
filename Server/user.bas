Attribute VB_Name = "user"
Option Explicit

Public Const STR_DK2SAPBC = "DK2SAPBC"
Public Const STR_DK2SAPYWZG44 = "DK2SAPYWZG44"
Public Const STR_DK2SAPYWZB09 = "DK2SAPYWZB09"
Public Const STR_DK2SAPYWZB35 = "DK2SAPYWZB35"
Public Const STR_DK2SAPYW = "DK2SAPYW"

Public Function is_system(str_user As String) As Boolean
    Select Case str_user
        Case STR_DK2SAPBC, STR_DK2SAPYWZG44, _
                STR_DK2SAPYWZB09, STR_DK2SAPYWZB35, _
                STR_DK2SAPYW
            is_system = True
        Case Else
            is_system = False
    End Select
End Function
