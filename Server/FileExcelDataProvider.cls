VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileExcelDataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const str_module As String = "FileExcelDataProvider"

Public str_id As String

Public str_path As String
Public str_file_name As String
Public str_source_type As String
'Public STR_FILE_APPENDIX As String

Public BOOL_EXTERNAL_DATA_FILE_VISIBILITY As Boolean

' local data
Public STR_WS_NAME As String
Public STR_DATA_FIRST_CELL As String

' info
Public obj_info As FileExcelDataProviderInfo

' status bar
Public STR_STATUS_BAR_PREFIX As String
Public STR_STATUS_BAR_PREFIX_LOADING As String
Public STR_STATUS_BAR_PREFIX_LOADING_FINISHED As String

Public wb As Workbook
Public rg_last_row As Range
'Public str_file_path As String
Public str_tmpl_file_path As String
Public bool_save_mode_on As Boolean
Public bool_visibility_mode As Boolean
Public bool_clear_after_open_status As Boolean
Public bool_clear_before_close_status As Boolean

'Public bool_record_as_processed_on As Boolean
Public str_completeness_status As String

' listeners must implement:
  ' load_record rg_record:Range
Public col_listeners As Collection

'Public Property Get str_id() As String
'    str_id = str_file_name
'End Property

Public Property Get file_path() As String
    file_path = str_path & str_file_name
End Property

Public Function set_clear_data_after_open_status(bool_yes_no As Boolean)
    'ws.UsedRange.Offset(1).clear
    'wb.Worksheets(STR_WS_NAME).UsedRange.Offset(1).clear
    bool_clear_after_open_status = bool_yes_no
End Function

Public Function set_clear_data_before_close_status(bool_yes_no As Boolean)
    'ws.UsedRange.Offset(1).clear
    'wb.Worksheets(STR_WS_NAME).UsedRange.Offset(1).clear
    bool_clear_before_close_status = bool_yes_no
End Function

Public Function open_data()
    'If file_path = "" Then
    If str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_INTERNAL Then
        Set wb = ThisWorkbook
    Else
        On Error GoTo ERR_FILE_NOT_FOUND
        Set wb = util_file.open_wb(file_path, Not bool_save_mode_on, is_visible:=bool_visibility_mode)
        On Error GoTo 0
    End If
    
    If bool_clear_after_open_status Then
        wb.Worksheets(STR_WS_NAME).UsedRange.Offset(1).clear
    End If
    
    Set rg_last_row = wb.Worksheets(STR_WS_NAME).Cells(wb.Worksheets(STR_WS_NAME).Range("A:A").CountLarge, 1)
    Exit Function
ERR_FILE_NOT_FOUND:
    If bool_save_mode_on Then
        Set wb = util_file.open_wb(str_tmpl_file_path, Not bool_save_mode_on, is_visible:=bool_visibility_mode)
        wb.SaveAs file_path ', WriteResPassword:=STR_PASSWD
        Resume Next
    Else
        Err.raise Err.Number, Err.source, Err.description
    End If
End Function

Public Function close_data()
    If bool_clear_before_close_status Then
        wb.Worksheets(STR_WS_NAME).UsedRange.Offset(1).clear
    End If

    If Not wb Is ThisWorkbook Then
        Windows(wb.name).Visible = True
        wb.Close saveChanges:=bool_save_mode_on
    End If
    
    Set wb = Nothing
End Function

Public Function load_data()
    Dim obj_listener As Object
    Dim errorMsg As MSG

    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING

    open_data
    ' init
    Set rg_record = wb.Worksheets(STR_WS_NAME).Range(STR_DATA_FIRST_CELL)
    
    Do While rg_record.Value <> ""
        DoEvents
                                        
'        If rg_record.Value = "357020105500244851" Then
'            DoEvents
'        End If
                                        
        On Error GoTo ERR_PROCESS
        For Each obj_listener In col_listeners
            obj_listener.load_record rg_record
        Next
        On Error GoTo 0
        ' move to next record
        Set rg_record = rg_record.Offset(1)
    Loop

    ' clean up
    Set obj_listener = Nothing
    close_data
    
    Application.StatusBar = STR_STATUS_BAR_PREFIX & STR_STATUS_BAR_PREFIX_LOADING_FINISHED
    Exit Function
ERR_PROCESS:
    Set errorMsg = New MSG
    log4VBA.warn log4VBA.DEFAULT_DESTINATION, errorMsg.source(str_module, "load_data").text("Unexpected error occurred during loading data. Listener: " & obj_listener.str_link_id & ", Record first column: " & rg_record.Value)
    'hndl_log.log db_log.TYPE_WARN, str_module, "load_data", "Unexpected error occurred during loading data. Listener: " & obj_listener.str_link_id & ", Record first column: " & rg_record.Value
    Resume Next
End Function

Public Function next_data() As Range
    Set next_data = rg_last_row.End(xlUp).Offset(1)
End Function

Public Function add_listener(obj_listener As Object)
    col_listeners.add obj_listener
End Function

Private Sub Class_Initialize()
    STR_DATA_FIRST_CELL = "A2"

    str_source_type = new_const_excel_data_provider.STR_SOURCE_TYPE_INTERNAL
    
    bool_save_mode_on = False
    bool_visibility_mode = False
    bool_clear_after_open_status = False
    bool_clear_before_close_status = False
    
    'bool_record_as_processed_on = True
    
    Set col_listeners = New Collection
                            
    ' status bar
    STR_STATUS_BAR_PREFIX = "Process Factory->"
    STR_STATUS_BAR_PREFIX_LOADING = "Loading ..."
    STR_STATUS_BAR_PREFIX_LOADING_FINISHED = "Loading has finished"
End Sub
