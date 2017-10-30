VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public destination As String
Public Level As Integer
Public category As Variant

Public Function log(lvl As Integer, dest As String, message As MSG)
    If Not lvl < Me.Level Then
        If dest = Me.destination Then
            save message
        End If
    End If
End Function

Public Function save(message As MSG)
    Debug.Print Now & " " & message.source_ & " " & message.text_
End Function

Public Function init(str_name As String, lvl As Integer, str_dest As String)
    Name = str_name
    Level = lvl
    destination = str_dest
End Function
