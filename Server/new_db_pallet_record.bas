Attribute VB_Name = "new_db_pallet_record"
Option Explicit

Public Const INT_OFFSET_PALLET As Integer = 0
Public Const INT_OFFSET_MATERIAL As Integer = 1
Public Const INT_OFFSET_SOURCE_DESTINATION As Integer = 2
Public Const INT_OFFSET_DATE As Integer = 3
Public Const INT_OFFSET_BIN As Integer = 4
Public Const INT_OFFSET_PROCESS_MASTER_ID As Integer = 5
Public Const INT_OFFSET_PROCESS_STEP_ORDER As Integer = 6
Public Const INT_OFFSET_PROCESS_STATUS As Integer = 7

Public Const STR_TASK_TYPE_CREATE As String = "CREATE"
Public Const STR_TASK_TYPE_CREATE_SUPPLY As String = "CREATE_SUPPLY"
Public Const STR_TASK_TYPE_UPDATE As String = "UPDATE"
Public Const STR_TASK_TYPE_CLOSE As String = "CLOSE"
