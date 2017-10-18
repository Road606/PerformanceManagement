Attribute VB_Name = "db_log"
Option Explicit

Public Const TYPE_INFO As String = "INFO"
Public Const TYPE_WARN As String = "WARN"
Public Const TYPE_ERR As String = "ERROR"

Public INT_DATA_COL_OFFSET_DATETIME As Integer
Public INT_DATA_COL_OFFSET_TYPE As Integer
Public INT_DATA_COL_OFFSET_MODULE As Integer
Public INT_DATA_COL_OFFSET_FUNCTION As Integer
Public INT_DATA_COL_OFFSET_MESSAGE As Integer

Public Function init()
    INT_DATA_COL_OFFSET_DATETIME = 0
    INT_DATA_COL_OFFSET_TYPE = 1
    INT_DATA_COL_OFFSET_MODULE = 2
    INT_DATA_COL_OFFSET_FUNCTION = 3
    INT_DATA_COL_OFFSET_MESSAGE = 4
End Function
