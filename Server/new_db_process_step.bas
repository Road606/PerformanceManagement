Attribute VB_Name = "new_db_process_step"
Option Explicit

Public Const BYTE_NEW As Byte = 1
Public Const BYTE_OPEN As Byte = 2
Public Const BYTE_CLOSED As Byte = 4

Public Const STR_NEW As String = "NEW"
Public Const STR_OPEN As String = "OPEN"
Public Const STR_CLOSED As String = "CLOSED"

Public Const STR_ORDER_STATUS_BACK = "Step back"
Public Const STR_ORDER_STATUS_AHEAD = "Step ahead"
Public Const STR_ORDER_STATUS_NOT_FOUND = "Step not found"
Public Const STR_ORDER_STATUS_OK = "Ok"

Public Const BYTE_PROCESS_STATUS_CONTINUE As Byte = 1
Public Const BYTE_PROCESS_STATUS_STOP As Byte = 2
