VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDLDataProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public obj_finished As Object
Public obj_unfinished As Object
'
'Public Function save_single(obj_process As process)
'    obj_finished.save_single obj_process
'End Function
'
Public Function open_connection()
    obj_finished.open_connection
    obj_unfinished.open_connection
End Function

Public Function close_connection()
    obj_finished.close_connection
    obj_unfinished.close_connection
End Function
