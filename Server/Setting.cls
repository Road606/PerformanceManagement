VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' private class variables
'Private mwsSettings As Worksheet
Public mwsSettings As Worksheet
Private mrgSetting As Range
Private mbAllowEditing As Boolean

' private class constants
Private Const SETTINGS_WORKSHEET = "settings"
Private Const VALUE_OFFSET = 1
Private Const TYPE_OFFSET = 2
Private Const DESCRIPTION_OFFSET = 3
Private Const CHANGE_EVENT_OFFSET = 4

Private Const str_module As String = "Setting"

' Enumeration for the kinds of Setting types
Enum setSettingType
setPrivate = 0 ' don't display in user inteface
setReadOnly = 1 ' read-only setting
setReadWrite = 2 ' read-write setting
setReadProtectedWrite = 3 ' read-write with password
End Enum

' Setting Description
Property Get description() As String
    If mrgSetting Is Nothing Then
        description = ""
    Else
        description = mrgSetting.Offset(0, DESCRIPTION_OFFSET).Value
    End If
End Property

Property Let description(PropertyDescription As String)
    If mrgSetting Is Nothing Then
        UninitializedError
    Else
        ' mbAllowEditing is managed by the
        ' EditMode method.
        If mbAllowEditing Then
            mrgSetting.Offset(0, DESCRIPTION_OFFSET).Value = _
            PropertyDescription
        Else
            ReadOnlyError
        End If
    End If
End Property

' setting EventHandler - represents a procedure that
' gets called automatically when the setting's value
' changes
Property Get EventHandler() As String
    If mrgSetting Is Nothing Then
        EventHandler = ""
    Else
        EventHandler = mrgSetting.Offset(0, CHANGE_EVENT_OFFSET).Value
    End If
    End Property
    
Property Let EventHandler(EventHandlerProcedure As String)
    If mrgSetting Is Nothing Then
        UninitializedError
    Else
        ' mbAllowEditing is managed by the
        ' EditMode method.
        If mbAllowEditing Then
            mrgSetting.Offset(0, CHANGE_EVENT_OFFSET).Value = _
            EventHandlerProcedure
        Else
            ReadOnlyError
        End If
    End If
End Property

' The settings are ordered by row
' on the Settings worksheet. Because this worksheet
' includes one row for column headings, you can
' get the index of the setting by looking at
' the row of the setting and subtracting one.
Property Get Index() As Long
    If mrgSetting Is Nothing Then
        Index = -1
    Else
        Index = mrgSetting.Row - 1
    End If
End Property

Property Get Name() As String
    If mrgSetting Is Nothing Then
        Name = ""
    Else
        Name = mrgSetting.Value
    End If
End Property

Property Let Name(PropertyName As String)
    ' Name is implemented as a read-only property
    ' so you can create dependencies on Setting
    ' names in your code.
End Property

Property Get SettingType() As setSettingType
    If mrgSetting Is Nothing Then
        SettingType = -1
    Else
        SettingType = mrgSetting.Offset(0, TYPE_OFFSET)
    End If
End Property

Property Let SettingType(SettingType As setSettingType)
    If mrgSetting Is Nothing Then
        UninitializedError
    Else
        If mbAllowEditing Then
            mrgSetting.Offset(0, TYPE_OFFSET).Value = SettingType
        Else
            ReadOnlyError
        End If
    End If
End Property

Property Get Value() As Variant
    If mrgSetting Is Nothing Then
        Value = ""
    Else
        Value = mrgSetting.Offset(0, VALUE_OFFSET)
    End If
    End Property
    
Property Let Value(PropertyValue As Variant)
    If mrgSetting Is Nothing Then
        UninitializedError
    Else
        If mbAllowEditing Then
            ' ok, change the value
            mrgSetting.Offset(0, VALUE_OFFSET).Value = PropertyValue
            ' call any procedures specified by the setting
            ' in the event of a change
            ExecuteEventHandler
        Else
            ReadOnlyError
        End If
    End If
End Property

Public Function Delete() As Boolean
    Delete = False
    If mrgSetting Is Nothing Then
        UninitializedError
    Else
        If mbAllowEditing Then
            mrgSetting.EntireRow.Delete xlUp
            Set mrgSetting = Nothing
            Delete = True
        Else
            ReadOnlyError
        End If
    End If
End Function

Public Function ChangeEditMode(AllowEditing As Boolean, _
    Optional Password As Variant) As Boolean
    If AllowEditing Then
        Select Case Me.SettingType
            Case setSettingType.setPrivate
                ' Private settings are settings used
                ' for programatic purposes or otherwise
                ' that should not be displayed on any
                ' user interface but can be freely
                ' modified programmatically
                mbAllowEditing = True
            Case setSettingType.setReadOnly
                ' Settings that are not intended to
                ' be changed by users but are useful
                ' to know. Never allow EditMode on these.
                mbAllowEditing = False
            Case setSettingType.setReadWrite
                ' Settings that can be freely modified
                ' by the user
                mbAllowEditing = True
            Case setSettingType.setReadProtectedWrite
                ' Settings that can be read but only
                ' changed by users that know the password
                If IsMissing(Password) Then
                    mbAllowEditing = False
                Else
                    If ValidPassword(CStr(Password)) Then
                        mbAllowEditing = True
                    Else
                        mbAllowEditing = False
                    End If
                End If
            Case Else
                ' unknown setting type
                mbAllowEditing = False
        End Select
    Else
        mbAllowEditing = False
    End If
    ChangeEditMode = mbAllowEditing
End Function

Public Function GetSetting(SettingName As String) As Boolean
    Dim lRow As Integer
    Dim bFoundSetting As Boolean
    Set mrgSetting = Nothing
    bFoundSetting = False
    mbAllowEditing = False
    lRow = 2
    
    Do Until IsEmpty(mwsSettings.Cells(lRow, 1))
        If UCase(mwsSettings.Cells(lRow, 1).Value) = _
            UCase(SettingName) Then
            Set mrgSetting = mwsSettings.Cells(lRow, 1)
            bFoundSetting = True
            Exit Do
        End If
        lRow = lRow + 1
    Loop
    GetSetting = bFoundSetting
End Function

Private Sub UninitializedError()
    Err.raise vbObjectError + 101, "Setting Class", _
    "The setting has not been properly initialized. " & _
    "Use the GetSetting method to initialize the setting."
End Sub

Private Sub ReadOnlyError()
    Err.raise vbObjectError + 102, "Setting Class", _
    "The setting you are trying to change is " & _
    "either read-only, requires a password, or " & _
    "you have not put the object in edit mode " & _
    "using the EditMode method."
End Sub

Private Sub Class_Initialize()
    ' do not allow editing by default
    mbAllowEditing = False
End Sub

Private Sub Class_InitializeObsolete()
    ' do not allow editing by default
    mbAllowEditing = False
    ' need to point the mws worksheet variable to the
    ' Settings worksheet
    If WorksheetExists(ThisWorkbook, SETTINGS_WORKSHEET) Then
        Set mwsSettings = ThisWorkbook.Worksheets(SETTINGS_WORKSHEET)
    Else
        Set mwsSettings = Nothing
        Err.raise vbObjectError + 100, "Setting Class", _
        "The worksheet named " & SETTINGS_WORKSHEET & _
        " could not be located."
    End If
End Sub

' Determines if a given worksheet name exists in a workbook
Private Function WorksheetExists(wb As Workbook, sName As String) _
    As Boolean
    Dim s As String
    
    On Error GoTo WorksheetExistsErr
    
    s = wb.Worksheets(sName).Name
    WorksheetExists = True
    Exit Function
WorksheetExistsErr:
    WorksheetExists = False
End Function

' Validate password by comparing it against the value
' given by the Password setting on the Settings worksheet.
' Obviously, this assumes that the worksheet is managed such
' that it cannot be easily retrieved/discovered.
' WARNING: this provides only very basic security and should
' not be used to protect sensitive data.
Private Function ValidPassword(sPassword As String) As Boolean
    Dim oSetting As Setting
    Dim bValid As Boolean
    bValid = False
    Set oSetting = New Setting
    
    If oSetting.GetSetting("Password") Then
        If oSetting.Value = sPassword Then
            bValid = True
        Else
            bValid = False
        End If
    Else
        bValid = False
    End If
    
    Set oSetting = Nothing
    ValidPassword = bValid
End Function

Private Sub ExecuteEventHandler()
    On Error Resume Next
    
    ' make sure there is an Event Handler
    ' for the setting
    If Len(Me.EventHandler) <> 0 Then
    ' call the procedure specified
    ' by the EventHandler property
        Application.run Me.EventHandler
    End If
End Sub
