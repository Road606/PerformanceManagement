Attribute VB_Name = "app_error_old"
Option Explicit

Public Const LEVEL_ERR As Integer = 1
Public Const LEVEL_WARN As Integer = 2
Public Const LEVEL_INFO As Integer = 4

Public Const ERR_NUMBER = 1024

Function get_err_level(err_level As Integer) As Integer
    get_err_level = ERR_NUMBER + err_level
End Function

Function retrieve_err_level(err_level As Integer) As Integer
    retrieve_err_level = err_level - ERR_NUMBER
End Function


