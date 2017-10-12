VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    ' class constants
    Private Const SETTINGS_WORKSHEET = "settings"
    Private Const NAME_COLUMN = 1
    Private Const VALUE_COLUMN = 2
    Private Const str_module As String = "Settings"
    ' class variables
    Private wb As Workbook
    Private mwsSettings As Worksheet
    ' count of settings
    
    Public str_path As String
    
    
Property Get Count()
    Count = mwsSettings.Cells(65536, 1).End(xlUp).Row - 1
End Property

Public Function init()
    str_path = ThisWorkbook.Path & "\settings-local.xlsx"
End Function

Public Function open_data()
    Set wb = util_file.open_wb(str_path)
    On Error GoTo ERR_MISSING_WS
    Set mwsSettings = wb.Worksheets(SETTINGS_WORKSHEET)
    On Error GoTo 0
    
    Exit Function
ERR_MISSING_WS:
    Err.raise app_error.get_err_level(app_error.LEVEL_ERR), "Settings>open_data", _
        "Worksheet: " & SETTINGS_WORKSHEET & " is missing."
End Function

Public Function close_data()
    On Error GoTo ERR_CLOSE_FILE
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function
ERR_CLOSE_FILE:
    Set hndl_log.obj_log_record = New DBLogRecord
    hndl_log.obj_log_record.str_datetime = Now
    hndl_log.obj_log_record.str_type = db_log.TYPE_ERR
    hndl_log.obj_log_record.str_module = str_module
    hndl_log.obj_log_record.str_function = "close_data"
    hndl_log.obj_log_record.str_message = "An error occured during closing file " & str_path & ". Original error message>" & Err.description
    hndl_log.save_record
End Function

' adds a new setting. returns setting object
' associated with the new setting.
Public Function add(Name As String) As Setting
    Dim lRow As Long
    Dim oSetting As Setting
    ' make sure a setting with this name
    ' doesn't already exist
    'Set oSetting = Me.Item(Name)
    
    If Not SettingExists(Name) Then
        ' find the last used row and move down one row
        lRow = mwsSettings.Cells(65536, 1).End(xlUp).Row + 1
        ' add the name of the new setting
        mwsSettings.Cells(lRow, 1) = Name
        ' set a reference to it
        Set oSetting = Me.Item(Name)
    Else
        ' the item already exists
        Err.raise vbObjectError + 201, "Settings Class", _
        "A setting named " & Name & " already exists."
        Set oSetting = Nothing
    End If
    Set add = oSetting
End Function

' deletes ALL settings
Public Function Delete() As Boolean
    mwsSettings.Range(mwsSettings.Cells(2, 1), _
    mwsSettings.Cells(65536, 4)).ClearContents
    Delete = True
End Function

' retrieves a setting by index or name
' retrieves by index if Index is numeric
' retrieves by name if Index is not numeric
Public Function Item(Index As Variant) As Setting
    Dim lRow As Long
    Dim lFoundRow As Long
    Dim oSetting As Setting
    Dim sName As String
    Set oSetting = New Setting
    Set oSetting.mwsSettings = mwsSettings
    
    ' if Index is numeric then assume
    ' that we are looking by index
    ' if Index is not numeric then assume
    ' that we are looking by name
    If IsNumeric(Index) Then
        ' get the name of the setting associated with
        ' the index. Row of setting = Index + 1 (header row)
        sName = mwsSettings.Cells(Index + 1, 1).Value
        ' make sure we got a name rather than an empty
        ' cell
        If Len(sName) <> 0 Then
            ' set a reference to the setting
            If oSetting.GetSetting(sName) Then
            Set Item = oSetting
            Else
            Err.raise 9, "Settings Class", _
            "Subscript out of range."
            End If
        Else
            Err.raise 9, "Settings Class", _
            "Subscript out of range."
        End If
    Else
        If oSetting.GetSetting(CStr(Index)) Then
            Set Item = oSetting
        Else
            Err.raise 9, "Settings Class", _
            "Subscript out of range."
        End If
    End If
End Function

' performs a reverse-lookup. looks up a setting by value
' rather than by name.
Public Function ItemByValue(Value As Variant) As Setting
    Dim lRow As Long
    Dim oSetting As Setting
    Dim bFound As Boolean
    Set oSetting = New Setting
    Set oSetting.mwsSettings = mwsSettings
    bFound = False
    
    For lRow = 2 To mwsSettings.Cells(65536, 1).End(xlUp).Row
        If Value = mwsSettings.Cells(lRow, VALUE_COLUMN).Value Then
            If oSetting.GetSetting( _
                mwsSettings.Cells(lRow, NAME_COLUMN).Value) Then
                Set ItemByValue = oSetting
                Else
                Err.raise 9, "Settings Class", _
                "Subscript out of range."
            End If
            bFound = True
        Exit For
        End If
    Next
    
    If Not bFound Then
        Set ItemByValue = Nothing
        Err.raise 9, "Settings Class", _
        "Subscript out of range."
    End If
End Function

Private Sub Class_InitializeObsolete()
    ' need to point the mws worksheet variable to the
    ' Settings worksheet
    If WorksheetExists(ThisWorkbook, SETTINGS_WORKSHEET) Then
        Set mwsSettings = ThisWorkbook.Worksheets(SETTINGS_WORKSHEET)
    Else
        Set mwsSettings = Nothing
        Err.raise vbObjectError + 200, "Settings Class", _
        "The worksheet named " & SETTINGS_WORKSHEET & _
        " could not be located."
    End If
End Sub

' Determines if a given worksheet name exists in a workbook
Private Function WorksheetExists(wb As Workbook, _
    sName As String) As Boolean
    Dim s As String
    
    On Error GoTo bWorksheetExistsErr
    
    s = wb.Worksheets(sName).Name
    WorksheetExists = True
    Exit Function
    
bWorksheetExistsErr:
    WorksheetExists = False
End Function

Private Function SettingExists(SettingName As String) As Boolean
    Dim oSetting As Setting
    On Error GoTo ErrHandler
    
    Set oSetting = Me.Item(SettingName)
    SettingExists = True
    Set oSetting = Nothing
    Exit Function
    
ErrHandler:
    SettingExists = False
End Function
