VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public source_ As String
Public text_ As String
Public params As Variant


Public Function source(str_module As String, str_function As String) As MSG
    Me.source_ = str_module & log4VBA.DEFAULT_SEPARATOR & str_function
    Set source = Me
End Function

Public Function text(str_text As String) As MSG
    Me.text_ = str_text
    Set text = Me
End Function


Public Function addParams(var_params As Variant) As MSG
    Me.params = var_params
    Set addParams = Me
End Function
