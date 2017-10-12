VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Area"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public str_id As String
Public str_name As String

Public col_steps As Collection

Public Function add_step(obj_step As AreaStep)
    col_steps.add obj_step, obj_step.str_id
End Function

Private Sub Class_Initialize()
    Set col_steps = New Collection
End Sub
