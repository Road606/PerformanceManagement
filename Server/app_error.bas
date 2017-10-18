Attribute VB_Name = "app_error"
Option Explicit

Public Const LEVEL_ERR As Integer = 1
Public Const LEVEL_WARN As Integer = 2
Public Const LEVEL_INFO As Integer = 4

Public Const ERR_NUMBER = 1024

Public Function raise(int_err_num As Integer, str_source As String, str_msg As String)
    Err.raise get_err_num(int_err_num), str_source, str_msg
End Function

Function get_err_num(err_num As Integer) As Integer
    get_err_num = ERR_NUMBER + err_num
End Function

Function retrieve_err_num(err_num As Integer) As Integer
    retrieve_err_num = err_num - ERR_NUMBER
End Function

Function get_err_level(err_level As Integer) As Integer
    get_err_level = ERR_NUMBER + err_level
End Function

Function retrieve_err_level(err_level As Integer) As Integer
    retrieve_err_level = err_level - ERR_NUMBER
End Function
