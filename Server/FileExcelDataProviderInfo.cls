VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileExcelDataProviderInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_provider_id As String
Public str_type As String
Public str_loading_type As String

Public obj_period As Period

Public str_completeness_status As String ' contains constant from const_excel_data_provider
'Public bool_record_as_processed_on As Boolean
